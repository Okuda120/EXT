Imports Common
Imports CommonEXT
Imports Npgsql
Imports EXTZ
Imports System.Configuration
Imports System.Net
Imports System.Net.Mail
Imports System.Text

Public Class LogicEXTC0103

    '変数宣言
    Private sqlEXTC0102 As New SqlEXTC0102          'sqlクラス
    Private sqlEXTC0103 As New SqlEXTC0103          'sqlクラス
    Private DataEXTC0103 As New DataEXTC0103

    '印刷用
    '出力ファイル名：利用承認書
    Private Const CERTIFICATES_FILENAME As String = "利用確認書_{0}.xlsx"
    '出力ファイル名：利用明細書（合算）
    Private Const DETAILS_FILENAME_ALL As String = "付帯設備利用明細書_{0}.xlsx"
    '出力ファイル名：利用明細書（日別）
    Private Const DETAILS_FILENAME_ONEDAY As String = "付帯設備利用明細書_{0}_{1}.xlsx"

    '帳票フォーマット名
    Private Const FORMAT_CERTIFICATES As String = "利用確認書_スタジオ.xlsx"                 ' 利用承認書
    Private Const FORMAT_DETAILS As String = "付帯設備利用明細書_スタジオ.xlsx"              ' 利用明細書

    ''' <summary>
    ''' 各種明細情報登録／更新
    ''' </summary>
    ''' <returns>boolean エラーコード  true 正常終了　false 異常終了</returns>
    ''' <remarks>予約情報登録
    ''' <para>作成情報：2015/09/12 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function RegYoyakuDetail(ByRef dataEXTC0102 As DataEXTC0102) As Boolean

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
            sqlEXTC0103.registerYoyakuInfo(Cmd, Cn, dataEXTC0102)
            '明細更新
            '付帯
            Dim rowCnt As Integer = 0
            Dim row As DataRow
            Dim ht As Hashtable = dataEXTC0102.PropHtFutai
            '削除された曜日データの削除
            sqlEXTC0103.deleteFutai(Cmd, Cn, dataEXTC0102.PropStrYoyakuNo, dataEXTC0102.PropListRiyobi)
            '登録・更新
            For Each Key As String In ht.Keys
                Table = ht(Key)
                row = Table.Rows(0)
                sqlEXTC0103.registerFutai(Cmd, Cn, row, dataEXTC0102.PropStrYoyakuNo)
            Next
            '付帯明細
            sqlEXTC0103.deleteFutaiDetail(Cmd, Cn, dataEXTC0102.PropStrYoyakuNo)
            ht = dataEXTC0102.PropHtFutaiDetail
            For Each Key As String In ht.Keys
                rowCnt = 0
                Table = ht(Key)
                Do While Table.Rows.Count > rowCnt
                    row = Table.Rows(rowCnt)
                    sqlEXTC0103.registerFutaiDetail(Cmd, Cn, row, dataEXTC0102.PropStrYoyakuNo)
                    rowCnt = rowCnt + 1
                Loop
            Next
            '請求/入金/EXAS入金(billpay_tbl)
            sqlEXTC0103.deleteBillPay(Cmd, Cn, dataEXTC0102.PropStrYoyakuNo)
            rowCnt = 0
            Table = dataEXTC0102.PropDsBillReq.Tables("BILLPAY_TBL")
            Dim innerTable As New DataTable
            Dim innerRow As DataRow
            Dim innerRowCnt As Integer = 0
            Do While Table.Rows.Count > rowCnt
                row = Table.Rows(rowCnt)
                '請求依頼番号の発番
                If IsDBNull(row("seikyu_irai_no")) And row("seikyu_input_flg") = "1" Then
                    row("seikyu_irai_no") = sqlEXTC0103.nextValSeikyuIraiNo(Cmd, Cn)
                End If
                '請求依頼更新
                sqlEXTC0103.registerBillPay(Cmd, Cn, row, dataEXTC0102.PropStrYoyakuNo)
                'EXASプロジェクト設定(project_tbl)
                'EXASプロジェクト設定付帯設備(project_tbl)
                'If row("seikyu_input_flg") = "1" Then                                                ' 2016.02.03 UPD h.hagiwara
                If row("seikyu_input_flg") = "1" And row("get_seikyu_input_flg") <> "1" Then          ' 2016.02.03 UPD h.hagiwara
                    Dim seikyuKey As String = row("seikyu_irai_no")
                    If SEIKYU_NAIYOU_RIYO = row("seikyu_naiyo") Or SEIKYU_NAIYOU_RIYOFUTAI = row("seikyu_naiyo") Then
                        If dataEXTC0102.PropHtExasPro.ContainsKey(seikyuKey) = True Then
                            innerTable = dataEXTC0102.PropHtExasPro(seikyuKey)
                        Else
                            innerTable = dataEXTC0102.PropHtExasPro(rowCnt.ToString)
                        End If
                        innerRowCnt = 0
                        sqlEXTC0103.deleteExasProject(Cmd, Cn, dataEXTC0102.PropStrYoyakuNo, row("seikyu_irai_no"), innerTable)
                        Do While innerTable.Rows.Count > innerRowCnt
                            innerRow = innerTable.Rows(innerRowCnt)
                            innerRow("seikyu_irai_no") = seikyuKey
                            innerRow("seikyu_irai_seq") = row("seq")
                            sqlEXTC0103.registerExasProject(Cmd, Cn, dataEXTC0102.PropStrYoyakuNo, innerRow)
                            innerRowCnt = innerRowCnt + 1
                        Loop
                    End If
                    If SEIKYU_NAIYOU_FUTAI = row("seikyu_naiyo") Or SEIKYU_NAIYOU_RIYOFUTAI = row("seikyu_naiyo") Then
                        If dataEXTC0102.PropHtExasProFutai.ContainsKey(seikyuKey) = True Then
                            innerTable = dataEXTC0102.PropHtExasProFutai(seikyuKey)
                        Else
                            innerTable = dataEXTC0102.PropHtExasProFutai(rowCnt.ToString)
                        End If
                        innerRowCnt = 0

                        If SEIKYU_NAIYOU_FUTAI = row("seikyu_naiyo") Then                       ' 2016.01.15 ADD  h.hagiwara
                            sqlEXTC0103.deleteExasProject(Cmd, Cn, dataEXTC0102.PropStrYoyakuNo, row("seikyu_irai_no"), innerTable)
                        End If                                                                  ' 2016.01.15 ADD  h.hagiwara
                        Do While innerTable.Rows.Count > innerRowCnt
                            innerRow = innerTable.Rows(innerRowCnt)
                            innerRow("seikyu_irai_no") = seikyuKey
                            innerRow("seikyu_irai_seq") = row("seq")
                            sqlEXTC0103.registerExasProject(Cmd, Cn, dataEXTC0102.PropStrYoyakuNo, innerRow)
                            innerRowCnt = innerRowCnt + 1
                        Loop
                    End If
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
    ''' 請求取得
    ''' </summary>
    ''' <param name="dataEXTC0102"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetBillReq(ByRef dataEXTC0102 As DataEXTC0102) As Boolean
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
            If sqlEXTC0103.selectBillReq(Adapter, Cn, dataEXTC0102) = False Then
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
    ''' <param name="dataEXTC0102"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetNyukin(ByRef dataEXTC0102 As DataEXTC0102) As Boolean
        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)
        Dim Tx As NpgsqlTransaction = Nothing
        Dim Adapter As New NpgsqlDataAdapter

        Dim Table As New DataTable()

        '2016.09.21 m.hayabuchi add start 不具合対応（入金取消後、入金リンクデータ枠のデータが残る）
        'PropDtExasNyukinを初期化 
        If Not dataEXTC0102.PropDtExasNyukin Is Nothing Then
            dataEXTC0102.PropDtExasNyukin.Dispose()
        End If
        dataEXTC0102.PropDtExasNyukin = Nothing
        '2016.09.21 m.hayabuchi add end 不具合対応

        Try
            Cn.Open()
            'SELECT用SQLCommandを作成
            Dim dtBillPay As DataTable = dataEXTC0102.PropDsBillReq.Tables("BILLPAY_TBL")
            Dim index As Integer = 0
            Dim row As DataRow
            Do While index < dtBillPay.Rows.Count
                row = dtBillPay.Rows(index)
                If row("nyukin_link_no") Is DBNull.Value = False Then
                    sqlEXTC0103.selectNyukin(Adapter, Cn, dataEXTC0102, row("nyukin_link_no"), "", index)
                ElseIf row("seikyu_irai_no") Is DBNull.Value = False Then
                    sqlEXTC0103.selectNyukin(Adapter, Cn, dataEXTC0102, "", row("seikyu_irai_no"), index)
                Else
                    sqlEXTC0103.selectNyukin(Adapter, Cn, dataEXTC0102, "", "", index)
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
    ''' 確認依頼確認記録情報取得
    ''' </summary>
    ''' <param name="dataEXTC0102"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetApproval(ByRef dataEXTC0102 As DataEXTC0102) As Boolean
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
            If sqlEXTC0103.selectApprovalReq(Adapter, Cn, dataEXTC0102) = False Then
                '異常終了
                Return False
            End If
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "確認依頼情報取得", Nothing, Adapter.SelectCommand)
            If sqlEXTC0103.selectApprovalRes(Adapter, Cn, dataEXTC0102) = False Then
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
    ''' 付帯情報取得
    ''' </summary>
    ''' <param name="dataEXTC0102"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetFutai(ByRef dataEXTC0102 As DataEXTC0102) As Boolean
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
            If sqlEXTC0103.selectFutaiHeadder(Adapter, Cn, dataEXTC0102) = False Then
                '異常終了
                Return False
            End If
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "付帯情報取得", Nothing, Adapter.SelectCommand)
            If sqlEXTC0103.selectFutaiDetail(Adapter, Cn, dataEXTC0102) = False Then
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
    ''' <param name="dataEXTC0102"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function UpFutai(ByRef dataEXTC0102 As DataEXTC0102) As Boolean
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
            If sqlEXTC0103.selectFutaiHeadderUp(Adapter, Cn, dataEXTC0102) = False Then
                '異常終了
                Return False
            End If
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "付帯情報取得", Nothing, Adapter.SelectCommand)
            If sqlEXTC0103.selectFutaiDetailUp(Adapter, Cn, dataEXTC0102) = False Then
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
    ''' <param name="dataEXTC0102"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function DelFutai(ByRef dataEXTC0102 As DataEXTC0102) As Boolean
        Dim htHeader As Hashtable = dataEXTC0102.PropHtFutai
        Dim htDetail As Hashtable = dataEXTC0102.PropHtFutaiDetail
        Dim exist As Boolean = False
        For Each Key As String In htHeader.Keys
            For Each dataRiyobi As CommonDataRiyobi In dataEXTC0102.PropListRiyobi
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
    ''' <param name="dataEXTC0102"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetExasRiyoryo(ByRef dataEXTC0102 As DataEXTC0102) As Boolean
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
            If sqlEXTC0103.selectExasRiyoryo(Adapter, Cn, dataEXTC0102) = False Then
                '異常終了
                Return False
            End If
            If sqlEXTC0103.selectExasFutai(Adapter, Cn, dataEXTC0102) = False Then
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
    ''' <param name="dataEXTC0102"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CalcExasRiyoryo(ByRef dataEXTC0102 As DataEXTC0102) As Boolean
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
            If sqlEXTC0103.selectExasRiyoryoUp(Adapter, Cn, dataEXTC0102) = False Then
                '異常終了
                Return False
            End If
            If sqlEXTC0103.selectExasFutaiUp(Adapter, Cn, dataEXTC0102) = False Then
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
    Public Function RegRirekiKiroku(ByRef dataEXTC0102 As DataEXTC0102) As Boolean

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
            Table = dataEXTC0102.PropDsApproval.Tables("IRAI_RIREKI_TBL")
            rowCnt = 0
            Do While Table.Rows.Count > rowCnt
                row = Table.Rows(rowCnt)
                row("yoyaku_no") = dataEXTC0102.PropStrYoyakuNo
                If IsDBNull(row("seq")) = True Then
                    Dim body As String = String.Format(CommonDeclareEXTC.MAIL_SHONIN_IRAI_BODY, dataEXTC0102.PropStrYoyakuNo, dataEXTC0102.PropStrShutsuenNm, dataEXTC0102.PropStrRiyoNm, GetRiyobi(dataEXTC0102, True), GetRiyobi(dataEXTC0102, False))
                    'sendMail(CommonDeclareEXTC.MAIL_SHONIN_IRAI_SUBJECT, body, False)       ' 2015.11.20 UPD h.hagiwara
                    sendMail(CommonDeclareEXTC.MAIL_SHONIN_IRAI_SUBJECT, body, True, dataEXTC0102.PropStrYoyakuNo)         ' 2015.11.20 UPD h.hagiwara
                    sqlEXTC0103.registerIraiRireki(Cmd, Cn, row)
                End If
                rowCnt = rowCnt + 1
            Loop
            Table = dataEXTC0102.PropDsApproval.Tables("CHECK_RIREKI_TBL")
            rowCnt = 0
            Do While Table.Rows.Count > rowCnt
                row = Table.Rows(rowCnt)
                row("yoyaku_no") = dataEXTC0102.PropStrYoyakuNo
                If IsDBNull(row("seq")) = True Then
                    Dim body As String = String.Format(CommonDeclareEXTC.MAIL_SHONIN_KAKUNIN_BODY, dataEXTC0102.PropStrYoyakuNo, dataEXTC0102.PropStrShutsuenNm, dataEXTC0102.PropStrRiyoNm, GetRiyobi(dataEXTC0102, True), GetRiyobi(dataEXTC0102, False))
                    'sendMail(CommonDeclareEXTC.MAIL_SHONIN_KAKUNIN_SUBJECT, body, True)      ' 2015.11.20 UPD h.hagiwara
                    sendMail(CommonDeclareEXTC.MAIL_SHONIN_KAKUNIN_SUBJECT, body, False, dataEXTC0102.PropStrYoyakuNo)      ' 2015.11.20 UPD h.hagiwara
                    sqlEXTC0103.registerCheckRireki(Cmd, Cn, row)
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
    ''' <param name="dataEXTC0102"></param>
    ''' <remarks></remarks>
    Public Function GetRiyobi(ByRef dataEXTC0102 As DataEXTC0102, ByVal isFrom As Boolean) As String
        Dim minDate As DateTime
        Dim maxDate As DateTime
        Dim isFirst As Boolean = True
        For Each dataRiyobi As CommonDataRiyobi In dataEXTC0102.PropListRiyobi
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
        If Mailsend(strMailTo, strMailCc, Nothing, StrSubject, StrBody, "", "") = False Then
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
            If sqlEXTC0103.selecMailTo(Adapter, Cn, IsShonin, Yoyakuno) = False Then
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
    ''' <param name="dataEXTC0102"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetRiyoKinTotal(ByVal dataEXTC0102 As DataEXTC0102) As Long
        Dim riyoKin As Long = 0
        For Each riyobi As CommonDataRiyobi In dataEXTC0102.PropListRiyobi
            riyoKin = riyoKin + riyobi.PropIntRiyoKin
        Next
        Return riyoKin
    End Function
    'Public Function GetRiyoKinTotal(ByVal dataEXTC0102 As DataEXTC0102) As Integer
    '    Dim riyoKin As Integer = 0
    '    For Each riyobi As CommonDataRiyobi In dataEXTC0102.PropListRiyobi
    '        riyoKin = riyoKin + riyobi.PropIntRiyoKin
    '    Next
    '    Return riyoKin
    'End Function

    ' 2015.12.21 UPD START↓ h.hagiwara
    ' ''' <summary>
    ' ''' 初期表示用の請求額(付帯料金)を取得する
    ' ''' </summary>
    ' ''' <param name="dataEXTC0102"></param>
    ' ''' <returns></returns>
    ' ''' <remarks></remarks>
    'Public Function GetFutairyo(ByVal dataEXTC0102 As DataEXTC0102) As Integer
    '    Dim futaiKin As Integer = 0
    '    Dim table As DataTable
    '    Dim row As DataRow
    '    Dim ht As Hashtable = dataEXTC0102.PropHtFutai
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
    ''' <param name="dataEXTC0102"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetFutairyo(ByVal dataEXTC0102 As DataEXTC0102) As Long
        Dim futaiKin As Long = 0
        Dim table As DataTable
        Dim row As DataRow
        Dim ht As Hashtable = dataEXTC0102.PropHtFutai
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
    ' ''' <param name="dataEXTC0102"></param>
    ' ''' <returns></returns>
    ' ''' <para>作成情報：2015.11.11 h.hagiwara</para>
    ' ''' <remarks></remarks>
    'Public Function GetFutairyoTax(ByVal dataEXTC0102 As DataEXTC0102) As Integer
    '    Dim futaiKintax As Integer = 0
    '    Dim table As DataTable
    '    Dim row As DataRow
    '    Dim ht As Hashtable = dataEXTC0102.PropHtFutai
    '    If dataEXTC0102.PropStrKashiKind = CommonDeclareEXT.KASHI_KIND_HOUSE Then
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
    ''' <param name="dataEXTC0102"></param>
    ''' <returns></returns>
    ''' <para>作成情報：2015.11.11 h.hagiwara</para>
    ''' <remarks></remarks>
    Public Function GetFutairyoTax(ByVal dataEXTC0102 As DataEXTC0102) As Long
        Dim futaiKintax As Long = 0
        Dim table As DataTable
        Dim row As DataRow
        Dim ht As Hashtable = dataEXTC0102.PropHtFutai
        If dataEXTC0102.PropStrKashiKind = CommonDeclareEXT.KASHI_KIND_HOUSE Then
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
    ''' <param name="dataEXTC0102"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetTaxfirstRiyobi(ByVal dataEXTC0102 As DataEXTC0102) As Double
        Dim tax As Double = 0
        If dataEXTC0102.PropStrKashiKind = KASHI_KIND_HOUSE Then
            Return 0
        End If
        For Each riyobi As CommonDataRiyobi In dataEXTC0102.PropListRiyobi
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
            If sqlEXTC0103.selectTax(Adapter, Cn, Riyobi) = False Then
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
    ''' <param name="dataEXTC0102"></param>
    ''' <param name="title1"></param>
    ''' <param name="seikyuNaiyo"></param>
    ''' <remarks></remarks>
    Public Sub SetSeikyuTitle(ByRef dataEXTC0102 As DataEXTC0102, _
                              ByRef title1 As String, _
                              ByVal seikyuNaiyo As String)
        Dim minDate As DateTime
        Dim maxDate As DateTime
        Dim isFirst As Boolean = True
        For Each dataRiyobi As CommonDataRiyobi In dataEXTC0102.PropListRiyobi
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
            sqlEXTC0103.updateCancellSeikyu(Cmd, Cn, row)
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
            sqlEXTC0103.updateCancellNyukin(Cmd, Cn, row)
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
    ''' 付帯設備利用明細書出力メイン処理
    ''' </summary>
    ''' <param name="dataEXTC0103">[IN]正式予約登録／詳細(スタジオ)画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>付帯設備利用明細書を出力する
    ''' <para>作成情報：2015/09/02 d.sonoda
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function OutputUseDetailsMain(ByVal dataEXTC0103 As DataEXTC0103) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        If dataEXTC0103.PropStrClickedIncident IsNot Nothing And _
            dataEXTC0103.PropStrClickedIncident <> "" Then
            dataEXTC0103.PropStrClickedIncident = DateTime.Parse(dataEXTC0103.PropStrClickedIncident).ToString("yyyy/MM/dd")
            '出力対象の日付がある場合、日別のデータを取得
            If GetUseDetailsOneDay(dataEXTC0103) = False Then
                Return False
            End If
        Else
            'ない場合、合算のデータを取得
            If GetUseDetailsAll(dataEXTC0103) = False Then
                Return False
            End If
        End If

        '付帯設備利用明細書を編集し、出力
        If OutputUseDetails(dataEXTC0103) = False Then
            Return False
        End If

        '終了ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

        '正常終了
        Return True

    End Function

    ''' <summary>
    ''' 付帯設備利用明細書出力処理
    ''' </summary>
    ''' <param name="dataEXTC0103">[IN]正式予約登録／詳細(スタジオ)画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>付帯設備利用明細書を出力する
    ''' <para>作成情報：2015/09/02 d.sonoda
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function OutputUseDetails(ByVal dataEXTC0103 As DataEXTC0103) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '編集用スプレッドシート
        Dim vwOutputSheet As New FarPoint.Win.Spread.FpSpread
        Dim downloadPath As String      'ダウンロードパス

        Try
            '合算フラグ設定
            If dataEXTC0103.PropStrClickedIncident IsNot Nothing And _
                dataEXTC0103.PropStrClickedIncident IsNot "" Then
                dataEXTC0103.PropBlnSumFlg = False
            Else
                '日付がない場合はTrue（合算）を設定
                dataEXTC0103.PropBlnSumFlg = True
            End If

            'フォーマット読み込み
            downloadPath = ConfigurationManager.AppSettings("formatPath") & FORMAT_DETAILS
            'vwOutputSheet.OpenExcel("Format\付帯設備利用明細書_フォーマット.xlsx")
            vwOutputSheet.OpenExcel(downloadPath)
            vwOutputSheet.ActiveSheet.ColumnCount = 12 '列数
            vwOutputSheet.ActiveSheet.RowCount = 63    '行数
            If dataEXTC0103.PropBlnSumFlg Then
                '合算の場合
                '付帯設備利用明細（合算）出力
                If SetUseDetailsDataAll(dataEXTC0103, vwOutputSheet.ActiveSheet) = False Then
                    Return False
                End If
            Else
                '日別の場合
                '付帯設備利用明細（日別）出力
                If SetUseDetailsDataOneDay(dataEXTC0103, vwOutputSheet.ActiveSheet) = False Then
                    Return False
                End If
            End If

            ' ファイル保存ダイアログ起動
            Dim fn As String = ""
            Using sfd As New SaveFileDialog()
                sfd.Filter = "利用明細書(*.xlsx)|*.xlsx"
                If dataEXTC0103.PropBlnSumFlg Then
                    sfd.FileName = String.Format(DETAILS_FILENAME_ALL, dataEXTC0103.PropStrReserveNo)
                Else
                    sfd.FileName = String.Format(DETAILS_FILENAME_ONEDAY, dataEXTC0103.PropStrReserveNo, _
                                                 DateTime.Parse(dataEXTC0103.PropStrClickedIncident).ToString("yyyyMMdd"))
                End If

                If sfd.ShowDialog() = DialogResult.OK Then
                    fn = sfd.FileName

                    ' Excelファイルに保存
                    vwOutputSheet.SaveExcel(fn, FarPoint.Excel.ExcelSaveFlags.UseOOXMLFormat)
                    MsgBox("Excel出力が完了しました。")
                End If
            End Using

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
    ''' 付帯設備利用明細書（合算）作成処理
    ''' </summary>
    ''' <param name="dataEXTC0103">[IN]正式予約登録／詳細(スタジオ)画面Dataクラス</param>
    ''' <param name="vwOutputSheet">[IN/OUT]付帯設備利用明細書スプレッドシート</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>付帯設備利用明細書（合算）を作成する
    ''' <para>作成情報：2015/09/02 d.sonoda
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function SetUseDetailsDataAll(ByVal dataEXTC0103 As DataEXTC0103, _
                                              ByRef vwOutputSheet As FarPoint.Win.Spread.SheetView) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try
            ' 2016.02.02 UPD START↓ h.hggiwara 内税のみの場合に利用料が出力されない不具合対応
            'If dataEXTC0103.PropDtUseDetails_Output IsNot Nothing And _
            '    dataEXTC0103.PropDtUseDetails_Output.Rows.Count > 0 Then

            '    With dataEXTC0103.PropDtUseDetails_Output
            '        'ヘッダ
            '        If dataEXTC0103.PropBlnGeneralFlg Then
            '            '一般貸出の場合
            '            vwOutputSheet.Cells("A2").Value = "利用明細書（合算）"
            '        Else
            '            '社内利用の場合
            '            vwOutputSheet.Cells("A2").Value = "社内利用明細書（合算）"
            '        End If

            '        vwOutputSheet.Cells("B1").Value = .Rows(0).Item(0) '予約番号
            '        vwOutputSheet.Cells("I1").Value = Now.ToString("yyyy年M月d日") '出力日
            '        vwOutputSheet.Cells("B3").Value = .Rows(0).Item(1) '利用者名
            '        vwOutputSheet.Cells("B4").Value = .Rows(0).Item(2) '利用日時
            '        vwOutputSheet.Cells("B5").Value = .Rows(0).Item(3) 'アーティスト名

            '        'スタジオ利用料
            '        vwOutputSheet.Cells("E8").Value = .Rows(0).Item(4) '延べ利用日数
            '        vwOutputSheet.Cells("F8").Value = .Rows(0).Item(5) '金額
            '        vwOutputSheet.Cells("G8").Value = .Rows(0).Item(6) '調整額

            '        '付帯設備
            '        For i As Integer = 0 To .Rows.Count - 1

            '            '税別の付帯設備
            '            vwOutputSheet.Cells(i + 8, 0).Value = i + 2 'インデックス
            '            vwOutputSheet.Cells(i + 8, 1).Value = .Rows(i).Item(7) '項目名
            '            vwOutputSheet.Cells(i + 8, 2).Value = .Rows(i).Item(8) '単位
            '            vwOutputSheet.Cells(i + 8, 3).Value = .Rows(i).Item(9) '単価
            '            vwOutputSheet.Cells(i + 8, 4).Value = .Rows(i).Item(10) '数量
            '            vwOutputSheet.Cells(i + 8, 5).Value = .Rows(i).Item(11) '金額
            '            vwOutputSheet.Cells(i + 8, 6).Value = .Rows(i).Item(12) '調整額

            '        Next

            '    End With
            '    If dataEXTC0103.PropBlnGeneralFlg Then
            '        vwOutputSheet.Cells("E48").Value = dataEXTC0103.PropIntTax / 100 '消費税
            '    End If

            'End If
            'ヘッダ
            If dataEXTC0103.PropBlnGeneralFlg Then
                '一般貸出の場合
                vwOutputSheet.Cells("A2").Value = "利用明細書（合算）"
            Else
                '社内利用の場合
                vwOutputSheet.Cells("A2").Value = "社内利用明細書（合算）"
            End If

            If dataEXTC0103.PropDtUseDetails_Output IsNot Nothing And _
                 dataEXTC0103.PropDtUseDetails_Output.Rows.Count > 0 Then

                With dataEXTC0103.PropDtUseDetails_Output

                    vwOutputSheet.Cells("B1").Value = .Rows(0).Item(0) '予約番号
                    vwOutputSheet.Cells("I1").Value = Now.ToString("yyyy年M月d日") '出力日
                    vwOutputSheet.Cells("B3").Value = .Rows(0).Item(1) '利用者名
                    vwOutputSheet.Cells("B4").Value = .Rows(0).Item(2) '利用日時
                    vwOutputSheet.Cells("B5").Value = .Rows(0).Item(3) 'アーティスト名

                    'スタジオ利用料
                    vwOutputSheet.Cells("E8").Value = .Rows(0).Item(4) '延べ利用日数
                    vwOutputSheet.Cells("F8").Value = .Rows(0).Item(5) '金額
                    vwOutputSheet.Cells("G8").Value = .Rows(0).Item(6) '調整額

                End With
            ElseIf dataEXTC0103.PropDtUseDetailsNoTax_Output IsNot Nothing And _
                dataEXTC0103.PropDtUseDetailsNoTax_Output.Rows.Count > 0 Then

                With dataEXTC0103.PropDtUseDetailsNoTax_Output

                    vwOutputSheet.Cells("B1").Value = .Rows(0).Item(0) '予約番号
                    vwOutputSheet.Cells("I1").Value = Now.ToString("yyyy年M月d日") '出力日
                    vwOutputSheet.Cells("B3").Value = .Rows(0).Item(1) '利用者名
                    vwOutputSheet.Cells("B4").Value = .Rows(0).Item(2) '利用日時
                    vwOutputSheet.Cells("B5").Value = .Rows(0).Item(3) 'アーティスト名

                    'スタジオ利用料
                    vwOutputSheet.Cells("E8").Value = .Rows(0).Item(4) '延べ利用日数
                    vwOutputSheet.Cells("F8").Value = .Rows(0).Item(5) '金額
                    vwOutputSheet.Cells("G8").Value = .Rows(0).Item(6) '調整額

                End With

            End If

            If dataEXTC0103.PropDtUseDetails_Output IsNot Nothing And _
                dataEXTC0103.PropDtUseDetails_Output.Rows.Count > 0 Then

                With dataEXTC0103.PropDtUseDetails_Output

                    '付帯設備
                    For i As Integer = 0 To .Rows.Count - 1

                        If IsDBNull(.Rows(i).Item(7)) Then                                 ' 2016.07.21 ADD h.hagiwara
                        Else                                                               ' 2016.07.21 ADD h.hagiwara
                            '税別の付帯設備
                            vwOutputSheet.Cells(i + 8, 0).Value = i + 2 'インデックス
                            vwOutputSheet.Cells(i + 8, 1).Value = .Rows(i).Item(7) '項目名
                            vwOutputSheet.Cells(i + 8, 2).Value = .Rows(i).Item(8) '単位
                            vwOutputSheet.Cells(i + 8, 3).Value = .Rows(i).Item(9) '単価
                            vwOutputSheet.Cells(i + 8, 4).Value = .Rows(i).Item(10) '数量
                            vwOutputSheet.Cells(i + 8, 5).Value = .Rows(i).Item(11) '金額
                            vwOutputSheet.Cells(i + 8, 6).Value = .Rows(i).Item(12) '調整額
                        End If                                                                 ' 2016.07.21 ADD h.hagiwara

                    Next

                End With
                If dataEXTC0103.PropBlnGeneralFlg Then
                    vwOutputSheet.Cells("E48").Value = dataEXTC0103.PropIntTax / 100 '消費税
                End If

            End If
            ' 2016.02.02 UPD END↑ h.hggiwara 内税のみの場合に利用料が出力されない不具合対応

            If dataEXTC0103.PropDtUseDetailsNoTax_Output IsNot Nothing And _
                dataEXTC0103.PropDtUseDetailsNoTax_Output.Rows.Count > 0 Then

                With dataEXTC0103.PropDtUseDetailsNoTax_Output
                    '付帯設備
                    For i As Integer = 0 To .Rows.Count - 1

                        If IsDBNull(.Rows(i).Item(7)) Then                                 ' 2016.07.21 ADD h.hagiwara
                        Else                                                               ' 2016.07.21 ADD h.hagiwara
                            '税込の付帯設備
                            vwOutputSheet.Cells(i + 51, 0).Value = i + 1 'インデックス
                            vwOutputSheet.Cells(i + 51, 1).Value = .Rows(i).Item(7) '項目名
                            vwOutputSheet.Cells(i + 51, 2).Value = .Rows(i).Item(8) '単位
                            vwOutputSheet.Cells(i + 51, 3).Value = .Rows(i).Item(9) '単価
                            vwOutputSheet.Cells(i + 51, 4).Value = .Rows(i).Item(10) '数量
                            vwOutputSheet.Cells(i + 51, 5).Value = .Rows(i).Item(11) '金額
                            vwOutputSheet.Cells(i + 51, 6).Value = .Rows(i).Item(12) '調整額
                        End If                                                                 ' 2016.07.21 ADD h.hagiwara

                    Next

                End With

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
    ''' 付帯設備利用明細書（日別）作成処理
    ''' </summary>
    ''' <param name="dataEXTC0103">[IN]正式予約登録／詳細(スタジオ)画面Dataクラス</param>
    ''' <param name="vwOutputSheet">[IN/OUT]付帯設備利用明細書スプレッドシート</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>付帯設備利用明細書（日別）を作成する
    ''' <para>作成情報：2015/09/02 d.sonoda
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function SetUseDetailsDataOneDay(ByVal dataEXTC0103 As DataEXTC0103, _
                                              ByRef vwOutputSheet As FarPoint.Win.Spread.SheetView) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try

            ' 2016.02.02 UPD START↓ h.hggiwara 内税のみの場合に利用料が出力されない不具合対応
            'If dataEXTC0103.PropDtUseDetails_Output IsNot Nothing And _
            '    dataEXTC0103.PropDtUseDetails_Output.Rows.Count > 0 Then

            '    With dataEXTC0103.PropDtUseDetails_Output
            '        '帳票へ書込み
            '        'ヘッダ
            '        If dataEXTC0103.PropBlnGeneralFlg Then
            '            '一般貸出の場合
            '            vwOutputSheet.Cells("A2").Value = "利用明細書"
            '        Else
            '            '社内利用の場合
            '            vwOutputSheet.Cells("A2").Value = "社内利用明細書"
            '        End If

            '        vwOutputSheet.Cells("B1").Value = .Rows(0).Item(0) '予約番号
            '        vwOutputSheet.Cells("I1").Value = Now.ToString("yyyy年M月d日") '出力日
            '        vwOutputSheet.Cells("B3").Value = .Rows(0).Item(1) '利用者名
            '        vwOutputSheet.Cells("B4").Value = .Rows(0).Item(2) '利用日時
            '        vwOutputSheet.Cells("B5").Value = .Rows(0).Item(3) 'アーティスト名

            '        'スタジオ利用料
            '        If .Rows(0).Item(4) = 0 Then
            '            vwOutputSheet.Cells("E8").Value = ""   '利用日数
            '        Else
            '            vwOutputSheet.Cells("E8").Value = 1    '利用日数
            '        End If
            '        vwOutputSheet.Cells("F8").Value = .Rows(0).Item(4) '金額
            '        vwOutputSheet.Cells("G8").Value = .Rows(0).Item(14) '調整額

            '        '付帯設備
            '        For i As Integer = 0 To .Rows.Count - 1

            '            '税別の付帯設備
            '            vwOutputSheet.Cells(i + 8, 0).Value = i + 2 'インデックス
            '            vwOutputSheet.Cells(i + 8, 1).Value = .Rows(i).Item(5) '項目名
            '            vwOutputSheet.Cells(i + 8, 2).Value = .Rows(i).Item(6) '単位
            '            vwOutputSheet.Cells(i + 8, 3).Value = .Rows(i).Item(7) '単価
            '            vwOutputSheet.Cells(i + 8, 4).Value = .Rows(i).Item(8) '数量
            '            vwOutputSheet.Cells(i + 8, 5).Value = .Rows(i).Item(9) '金額
            '            vwOutputSheet.Cells(i + 8, 6).Value = .Rows(i).Item(10) '調整額
            '            vwOutputSheet.Cells(i + 8, 8).Value = .Rows(i).Item(11) '備考

            '        Next

            '    End With

            '    If dataEXTC0103.PropBlnGeneralFlg Then
            '        vwOutputSheet.Cells("E48").Value = dataEXTC0103.PropIntTax / 100 '消費税
            '    End If

            'End If
            'ヘッダ
            If dataEXTC0103.PropBlnGeneralFlg Then
                '一般貸出の場合
                vwOutputSheet.Cells("A2").Value = "利用明細書"
            Else
                '社内利用の場合
                vwOutputSheet.Cells("A2").Value = "社内利用明細書"
            End If

            If dataEXTC0103.PropDtUseDetails_Output IsNot Nothing And _
               dataEXTC0103.PropDtUseDetails_Output.Rows.Count > 0 Then

                With dataEXTC0103.PropDtUseDetails_Output
                    '帳票へ書込み

                    vwOutputSheet.Cells("B1").Value = .Rows(0).Item(0) '予約番号
                    vwOutputSheet.Cells("I1").Value = Now.ToString("yyyy年M月d日") '出力日
                    vwOutputSheet.Cells("B3").Value = .Rows(0).Item(1) '利用者名
                    vwOutputSheet.Cells("B4").Value = .Rows(0).Item(2) '利用日時
                    vwOutputSheet.Cells("B5").Value = .Rows(0).Item(3) 'アーティスト名

                    'スタジオ利用料
                    If .Rows(0).Item(4) = 0 Then
                        vwOutputSheet.Cells("E8").Value = ""   '利用日数
                    Else
                        vwOutputSheet.Cells("E8").Value = 1    '利用日数
                    End If
                    vwOutputSheet.Cells("F8").Value = .Rows(0).Item(4) '金額
                    vwOutputSheet.Cells("G8").Value = .Rows(0).Item(14) '調整額


                End With

            ElseIf dataEXTC0103.PropDtUseDetailsNoTax_Output IsNot Nothing And _
                dataEXTC0103.PropDtUseDetailsNoTax_Output.Rows.Count > 0 Then

                With dataEXTC0103.PropDtUseDetailsNoTax_Output
                    '帳票へ書込み

                    vwOutputSheet.Cells("B1").Value = .Rows(0).Item(0) '予約番号
                    vwOutputSheet.Cells("I1").Value = Now.ToString("yyyy年M月d日") '出力日
                    vwOutputSheet.Cells("B3").Value = .Rows(0).Item(1) '利用者名
                    vwOutputSheet.Cells("B4").Value = .Rows(0).Item(2) '利用日時
                    vwOutputSheet.Cells("B5").Value = .Rows(0).Item(3) 'アーティスト名

                    'スタジオ利用料
                    If .Rows(0).Item(4) = 0 Then
                        vwOutputSheet.Cells("E8").Value = ""   '利用日数
                    Else
                        vwOutputSheet.Cells("E8").Value = 1    '利用日数
                    End If
                    vwOutputSheet.Cells("F8").Value = .Rows(0).Item(4) '金額
                    vwOutputSheet.Cells("G8").Value = .Rows(0).Item(14) '調整額

                End With

            End If

            If dataEXTC0103.PropDtUseDetails_Output IsNot Nothing And _
               dataEXTC0103.PropDtUseDetails_Output.Rows.Count > 0 Then

                With dataEXTC0103.PropDtUseDetails_Output
                    '付帯設備
                    For i As Integer = 0 To .Rows.Count - 1

                        If IsDBNull(.Rows(i).Item(5)) Then                                 ' 2016.07.21 ADD h.hagiwara
                        Else                                                               ' 2016.07.21 ADD h.hagiwara
                            '税別の付帯設備
                            vwOutputSheet.Cells(i + 8, 0).Value = i + 2 'インデックス
                            vwOutputSheet.Cells(i + 8, 1).Value = .Rows(i).Item(5) '項目名
                            vwOutputSheet.Cells(i + 8, 2).Value = .Rows(i).Item(6) '単位
                            vwOutputSheet.Cells(i + 8, 3).Value = .Rows(i).Item(7) '単価
                            vwOutputSheet.Cells(i + 8, 4).Value = .Rows(i).Item(8) '数量
                            vwOutputSheet.Cells(i + 8, 5).Value = .Rows(i).Item(9) '金額
                            vwOutputSheet.Cells(i + 8, 6).Value = .Rows(i).Item(10) '調整額
                            vwOutputSheet.Cells(i + 8, 8).Value = .Rows(i).Item(11) '備考
                        End If                                                             ' 2016.07.21 ADD h.hagiwara

                    Next

                End With

                If dataEXTC0103.PropBlnGeneralFlg Then
                    vwOutputSheet.Cells("E48").Value = dataEXTC0103.PropIntTax / 100 '消費税
                End If

            End If
            ' 2016.02.02 UPD END↑ h.hggiwara 内税のみの場合に利用料が出力されない不具合対応

            If dataEXTC0103.PropDtUseDetailsNoTax_Output IsNot Nothing And _
                dataEXTC0103.PropDtUseDetailsNoTax_Output.Rows.Count > 0 Then

                With dataEXTC0103.PropDtUseDetailsNoTax_Output
                    '付帯設備
                    For i As Integer = 0 To .Rows.Count - 1

                        If IsDBNull(.Rows(i).Item(5)) Then                                 ' 2016.07.21 ADD h.hagiwara
                        Else                                                               ' 2016.07.21 ADD h.hagiwara
                            '税込の付帯設備
                            vwOutputSheet.Cells(i + 51, 0).Value = i + 1 'インデックス
                            vwOutputSheet.Cells(i + 51, 1).Value = .Rows(i).Item(5) '項目名
                            vwOutputSheet.Cells(i + 51, 2).Value = .Rows(i).Item(6) '単位
                            vwOutputSheet.Cells(i + 51, 3).Value = .Rows(i).Item(7) '単価
                            vwOutputSheet.Cells(i + 51, 4).Value = .Rows(i).Item(8) '数量
                            vwOutputSheet.Cells(i + 51, 5).Value = .Rows(i).Item(9) '金額
                            vwOutputSheet.Cells(i + 51, 6).Value = .Rows(i).Item(10) '調整額
                            vwOutputSheet.Cells(i + 51, 8).Value = .Rows(i).Item(11) '備考
                        End If                                                             ' 2016.07.21 ADD h.hagiwara

                    Next

                End With

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
    ''' 付帯設備利用明細（合算）取得処理
    ''' </summary>
    ''' <param name="dataEXTC0103">[IN/OUT]正式予約登録/詳細画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>ＤＢより付帯設備利用明細（合算）を取得する
    ''' <para>作成情報：2015/08/14 d.sonoda
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function GetUseDetailsAll(ByRef dataEXTC0103 As DataEXTC0103) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim dtUseDetailsAll As New DataTable
        Dim Cn As New NpgsqlConnection(DbString)  'サーバーとクライアントをつなげる
        Dim Adapter As New NpgsqlDataAdapter

        Try
            'コネクションを開く
            Cn.Open()

            '税無フラグ
            dataEXTC0103.PropBlnNoTaxFlg = False
            '付帯設備利用明細（合算）取得SQLを設定
            If sqlEXTC0103.SetSelectUseDetailsAllSql(Adapter, Cn, dataEXTC0103) = False Then
                Return False
            End If

            '付帯設備利用明細（合算）を取得
            Adapter.Fill(dtUseDetailsAll)
            dataEXTC0103.PropDtUseDetails_Output = dtUseDetailsAll

            dtUseDetailsAll = New DataTable

            '税無フラグ
            dataEXTC0103.PropBlnNoTaxFlg = True
            '付帯設備利用明細（合算）取得SQLを設定
            If sqlEXTC0103.SetSelectUseDetailsAllSql(Adapter, Cn, dataEXTC0103) = False Then
                Return False
            End If

            '付帯設備利用明細（合算）を取得
            Adapter.Fill(dtUseDetailsAll)
            dataEXTC0103.PropDtUseDetailsNoTax_Output = dtUseDetailsAll

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
    ''' <param name="dataEXTC0103">[IN/OUT]正式予約登録/詳細画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>ＤＢより付帯設備利用明細（日別）を取得する
    ''' <para>作成情報：2015/08/14 d.sonoda
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function GetUseDetailsOneDay(ByRef dataEXTC0103 As DataEXTC0103) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim dtUseDetailsOneDay As New DataTable
        Dim Cn As New NpgsqlConnection(DbString)  'サーバーとクライアントをつなげる
        Dim Adapter As New NpgsqlDataAdapter

        Try
            'コネクションを開く
            Cn.Open()

            '税無フラグ
            dataEXTC0103.PropBlnNoTaxFlg = False
            '付帯設備利用明細（合算）取得SQLを設定
            If sqlEXTC0103.SetSelectUseDetailsOneDaySql(Adapter, Cn, dataEXTC0103) = False Then
                Return False
            End If

            '付帯設備利用明細（合算）を取得
            Adapter.Fill(dtUseDetailsOneDay)
            dataEXTC0103.PropDtUseDetails_Output = dtUseDetailsOneDay

            dtUseDetailsOneDay = New DataTable

            '税無フラグ
            dataEXTC0103.PropBlnNoTaxFlg = True
            '付帯設備利用明細（合算）取得SQLを設定
            If sqlEXTC0103.SetSelectUseDetailsOneDaySql(Adapter, Cn, dataEXTC0103) = False Then
                Return False
            End If

            '付帯設備利用明細（合算）を取得
            Adapter.Fill(dtUseDetailsOneDay)
            dataEXTC0103.PropDtUseDetailsNoTax_Output = dtUseDetailsOneDay

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
    ''' 利用確認書取得処理予定地
    ''' </summary>
    ''' <param name="dataEXTC0103">[IN/OUT]正式予約登録/詳細画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>ＤＢより利用確認書を取得する予定
    ''' <para>作成情報：2015/08/31 yu.satoh
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function OutputStudioData(ByRef dataEXTC0103 As DataEXTC0103) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try

            'SQL作成
            If SetSqlSelectData(dataEXTC0103) = False Then
                Return False
            End If

            '取得データをエクセルフォーマットにセット
            If SetPrintData(dataEXTC0103) = False Then
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
            puErrMsg = EXT_E001 & ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' 利用確認書取得SQL作成
    ''' </summary>
    ''' <param name="dataEXTC0103">[IN/OUT]正式予約登録/詳細画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>情報を取得するSQLを作成する。
    ''' <para>作成情報：2015/08/31 yu.satoh
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function SetSqlSelectData(ByRef dataEXTC0103 As DataEXTC0103) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)  'サーバーとクライアントをつなげる
        Dim Adapter As New NpgsqlDataAdapter
        Dim Table As New DataTable                'テーブル

        Try
            'コネクションを開く
            Cn.Open()

            'スタジオ予約表取得SQLの作成・設定
            If sqlEXTC0103.SetSelectStudioData(Adapter, Cn, dataEXTC0103) = False Then
                Return False
            End If

            'SQLログ
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "システム管理情報", Nothing, Adapter.SelectCommand)

            'データを取得
            Adapter.Fill(Table)
            '取得データをデータクラスへ保存
            dataEXTC0103.PropDtReportData = Table

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
            Table.Dispose()
            Adapter.Dispose()
            Cn.Dispose()
        End Try

    End Function


    ''' <summary>
    ''' 利用確認書出力処理
    ''' </summary>
    ''' <param name="dataEXTC0103">[IN]正式予約登録/詳細画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>利用確認書を編集し、印刷する
    ''' <para>作成情報：2015/08/10 d.sonoda
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function SetPrintData(ByVal dataEXTC0103 As DataEXTC0103) As Boolean
        '変数
        Dim intRiyoTimeH As Integer '利用時間
        Dim intRiyoTimeM As Double '利用時間分
        Dim downloadpath As String 'ダウンロードファイルパス


        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try
            With dataEXTC0103
                'フォーマット読み込み
                downloadPath = ConfigurationManager.AppSettings("formatPath") & FORMAT_CERTIFICATES

                'Excel読み込み
                .PropVwStudioSheet = New FarPoint.Win.Spread.FpSpread
                '.PropVwStudioSheet.OpenExcel("..\..\..\EXTA\Format\正式予約確認書_スタジオ.xlsx")
                .PropVwStudioSheet.OpenExcel(downloadPath)

                '結合されている場合左上のセルを参照する
                ' シート設定
                Dim vwAppSheet As FarPoint.Win.Spread.SheetView = .PropVwStudioSheet.Sheets(0) '利用承認書
                vwAppSheet.ColumnCount = 20 '列数
                vwAppSheet.RowCount = 40    '行数

                '付帯設備利用料
                If .PropDtReportData IsNot Nothing Then
                    'DataTableから一行取得
                    Dim dtSelectRow As DataRow = .PropDtReportData.Rows(0)
                    '利用時間計算用変数
                    If IsDBNull(dtSelectRow.Item(10)) And IsDBNull(dtSelectRow.Item(11)) Then
                        intRiyoTimeH = 0
                        intRiyoTimeM = 0
                    Else
                        Dim intStartTimeH As Integer = CInt(dtSelectRow.Item(10).ToString.Substring(0, 2))
                        Dim intStartTimeM As Integer = CInt(dtSelectRow.Item(10).ToString.Substring(2, 2))
                        Dim intEndTimeH As Integer = CInt(dtSelectRow.Item(11).ToString.Substring(0, 2))
                        Dim intEndTimeM As Integer = CInt(dtSelectRow.Item(11).ToString.Substring(2, 2))
                        '利用時間を算出
                        intRiyoTimeH = intEndTimeH - intStartTimeH
                        intRiyoTimeM = intEndTimeM - intStartTimeM
                    End If
                    '担当者、テストのため現在べたうち.本来はここにグローバル変数
                    .PropLoginUser = CommonEXT.PropComStrUserName

                    '利用確認書に値を設定
                    vwAppSheet.Cells("C3").Value = dtSelectRow.Item(0)   '承認番号(画面上予約番号)
                    vwAppSheet.Cells("A4").Value = dtSelectRow.Item(1)   '利用者名＋責任者名
                    vwAppSheet.Cells("D10").Value = dtSelectRow.Item(2) & _
                                                    dtSelectRow.Item(4)  '予約受付日 + 曜日  
                    vwAppSheet.Cells("H11").Value = dtSelectRow.Item(3) & _
                                                    dtSelectRow.Item(4)  '予約有効期限 + 曜日
                    vwAppSheet.Cells("F12").Value = dtSelectRow.Item(5)  'アーティスト名
                    vwAppSheet.Cells("A14").Value = dtSelectRow.Item(6)  '利用日時(SQLで曜日を追加)
                    vwAppSheet.Cells("J14").Value = dtSelectRow.Item(7)  '利用日数
                    vwAppSheet.Cells("L14").Value = dtSelectRow.Item(8) & vbCrLf & _
                                                    dtSelectRow.Item(9)  'スタジオ区分 + 利用形態
                    '利用時間帯 開始～終了,時間の間に「：」を入れる
                    If IsDBNull(dtSelectRow.Item(10)) And IsDBNull(dtSelectRow.Item(11)) Then
                    Else
                        vwAppSheet.Cells("M14").Value =
                                                        dtSelectRow.Item(10).ToString.Substring(0, 2) & _
                                                        "：" & dtSelectRow.Item(10).ToString.Substring(2, 2) & _
                                                         "～" & _
                                                        dtSelectRow.Item(11).ToString.Substring(0, 2) & _
                                                        "：" & dtSelectRow.Item(11).ToString.Substring(2, 2)
                    End If

                    '利用時間の分を計算
                    If intRiyoTimeM < 0 Then
                        intRiyoTimeH -= 1
                        intRiyoTimeM = (intRiyoTimeM + 60) / 6
                        '小数切り捨て
                        intRiyoTimeM = Math.Floor(intRiyoTimeM)
                    ElseIf intRiyoTimeM > 0 Then
                        intRiyoTimeM = intRiyoTimeM / 6
                        '小数切り捨て
                        intRiyoTimeM = Math.Floor(intRiyoTimeM)
                    End If

                    '利用時間,端数が出なければ出力は時間のみ
                    If intRiyoTimeM = 0 Then
                        vwAppSheet.Cells("P14").Value = intRiyoTimeH
                    Else
                        vwAppSheet.Cells("P14").Value = intRiyoTimeH & "." & intRiyoTimeM
                    End If

                    '出力日
                    vwAppSheet.Cells("K40").Value = dtSelectRow.Item(2).ToString.Substring(0, 5) & "  " & _
                                                    dtSelectRow.Item(2).ToString.Substring(5, 3) & "  " & _
                                                    dtSelectRow.Item(2).ToString.Substring(8, 3)
                    '出力者名
                    vwAppSheet.Cells("P40").Value = .PropLoginUser

                End If

                ' ファイル保存ダイアログ起動
                Dim fn As String = ""
                Using sfd As New SaveFileDialog()
                    sfd.Filter = "利用確認書(*.xlsx)|*.xlsx"
                    sfd.FileName = String.Format(CERTIFICATES_FILENAME, .PropLblYoyakuNo.Text)
                    If sfd.ShowDialog() = DialogResult.OK Then
                        fn = sfd.FileName

                        ' Excelファイルに保存
                        .PropVwStudioSheet.SaveExcel(fn, FarPoint.Excel.ExcelSaveFlags.UseOOXMLFormat)
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
            If sqlEXTC0103.GetHolidayMst(Adapter, Cn, CehckDay) = False Then
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
    ''' <paramref name="DataEXTC0103">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean エラーコード    True:正常  False:異常</returns>
    ''' <remarks>EXAS連携時に消費税率の設定
    ''' <para>作成情報：2015.11.11 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function GetNotaxflg(ByRef DataEXTC0103 As DataEXTC0103) As String

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
            If GetNotaxflgInf(Adapter, Cn, DataEXTC0103) = False Then
                strTaxlfg = "0"
            Else
                strTaxlfg = DataEXTC0103.PropStrNoTaxFlg
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
    ''' <param name="DataEXTC0103">[IN/OUT]登録画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>EXAS連携時に消費税率の設定を取得する
    ''' <para>作成情報：2015.11.11 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function GetNotaxflgInf(ByVal Adapter As NpgsqlDataAdapter, _
                               ByVal Cn As NpgsqlConnection, _
                               ByRef DataEXTC0103 As DataEXTC0103) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim dtmst As New DataTable

        Try
            '取得用SQLの作成・設定
            If sqlEXTC0103.GeSqlNotaxflg(Adapter, Cn, DataEXTC0103) = False Then
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
                    DataEXTC0103.PropStrNoTaxFlg = dtmst.Rows(0).Item(0)
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
    ''' <paramref name="dataEXTC0103">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean エラーコード    True:正常  False:異常</returns>
    ''' <remarks>相手先情報の取得
    ''' <para>作成情報：2015.11.11 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function GetAitesakiInf(ByRef dataEXTC0103 As DataEXTC0103) As Boolean

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
            If GetAitesaki(Adapter, Cn, dataEXTC0103) = False Then
                Return False
            Else
                strTaxlfg = dataEXTC0103.PropStrNoTaxFlg
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
    ''' <param name="dataEXTC0103">[IN/OUT]登録画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>EXAS連携時に消費税率の設定を取得する
    ''' <para>作成情報：2015.11.11 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function GetAitesaki(ByVal Adapter As NpgsqlDataAdapter, _
                               ByVal Cn As NpgsqlConnection, _
                               ByRef dataEXTC0103 As DataEXTC0103) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim dtmst As New DataTable

        Try
            '取得用SQLの作成・設定
            If sqlEXTC0103.GeSqlAitesaki(Adapter, Cn, dataEXTC0103) = False Then
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
                    dataEXTC0103.PropStrPostno = dtmst.Rows(0).Item(0)
                    dataEXTC0103.PropStrAddr1 = dtmst.Rows(0).Item(1)
                    dataEXTC0103.PropStrAddr2 = dtmst.Rows(0).Item(2)
                    dataEXTC0103.PropStrAitenm = dtmst.Rows(0).Item(3)
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
    ''' <para>作成情報：2015.12.10 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function RegYoyakuInfo(ByRef dataEXTC0102 As DataEXTC0102) As Boolean

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
            sqlEXTC0103.registerYoyakuInfo(Cmd, Cn, dataEXTC0102)

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
    Public Function GetBillpayFlg(ByRef dataEXTC0102 As DataEXTC0102) As Boolean
        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim dtBillPay As DataTable = dataEXTC0102.PropDsBillReq.Tables("BILLPAY_TBL")
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
    Public Function GetRyokinBairituMst(ByRef dataEXTC0103 As DataEXTC0103) As Boolean

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)  'サーバーとクライアントをつなげる
        Dim Adapter As New NpgsqlDataAdapter      'アダプタ

        Try
            'コネクションを開く
            Cn.Open()

            ' 料金マスタの取得
            If GetRyokinMst(Adapter, Cn, dataEXTC0103) = False Then
                Return False
            End If

            ' 倍率マスタの取得
            If GeBairituMst(Adapter, Cn, dataEXTC0103) = False Then
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
    ''' <param name="dataEXTC0103">[IN/OUT]登録画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>料金・倍率選択画面表示前に料金マスタを取得する
    ''' <para>作成情報：2015.12.16 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function GetRyokinMst(ByVal Adapter As NpgsqlDataAdapter, _
                               ByVal Cn As NpgsqlConnection, _
                               ByRef dataEXTC0103 As DataEXTC0103) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim dtmst As New DataTable

        Try
            '取得用SQLの作成・設定
            If sqlEXTC0103.GeSqlRyokinMst(Adapter, Cn, dataEXTC0103) = False Then
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
    ''' <param name="dataEXTC0103">[IN/OUT]登録画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>料金・倍率選択画面表示前に倍率マスタを取得する
    ''' <para>作成情報：2015.12.16 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function GeBairituMst(ByVal Adapter As NpgsqlDataAdapter, _
                               ByVal Cn As NpgsqlConnection, _
                               ByRef dataEXTC0103 As DataEXTC0103) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim dtmst As New DataTable

        Try
            '取得用SQLの作成・設定
            If sqlEXTC0103.GeSqlBairituMst(Adapter, Cn, dataEXTC0103) = False Then
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
    Public Function GetFutaiSetubiInfMst(ByRef dataEXTC0103 As DataEXTC0103) As Boolean

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)  'サーバーとクライアントをつなげる
        Dim Adapter As New NpgsqlDataAdapter      'アダプタ

        Try
            'コネクションを開く
            Cn.Open()

            ' 付帯設備マスタの取得
            If GetFutaiSetubiMst(Adapter, Cn, dataEXTC0103) = False Then
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
    ''' <param name="dataEXTC0103">[IN/OUT]登録画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>付帯設備設定画面表示前に付帯設備マスタを取得する
    ''' <para>作成情報：2015.12.16 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function GetFutaiSetubiMst(ByVal Adapter As NpgsqlDataAdapter, _
                               ByVal Cn As NpgsqlConnection, _
                               ByRef dataEXTC0103 As DataEXTC0103) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim dtmst As New DataTable

        Try
            '取得用SQLの作成・設定
            If sqlEXTC0103.GeSqlFutaiSetubiMst(Adapter, Cn, dataEXTC0103) = False Then
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
    ''' <para>作成情報：2016/11/4　m.hayabuchi
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function GetSekininshaMailTel(ByVal dataEXTC0102 As DataEXTC0102) As Boolean

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
            If sqlEXTC0103.GetSqlSekininshaMailTelData(Adapter, Cn, dataEXTC0102) = False Then
                Return False
            End If

            'SQLログ
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "責任者メールアドレス・携帯電話番号情報取得", Nothing, Adapter.SelectCommand)

            'データを取得
            Adapter.Fill(Table)

            '取得した項目をDataクラスに格納
            With dataEXTC0102
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
