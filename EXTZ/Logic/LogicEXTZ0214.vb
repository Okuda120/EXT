Imports Common
Imports CommonEXT
Imports System.Text
Imports System.IO
Imports System.Configuration

Public Class LogicEXTZ0214

    Private dataEXTZ0214 As New DataEXTZ0214          'Dataクラス
    Private commonLogic As New CommonLogic          '共通ロジッククラス
    Private commonLogicEXT As New CommonLogicEXT    'EXT共通ロジッククラス

    ''' <summary>
    ''' ＣＳＶファイル取り込み処理
    ''' <paramref name="dataEXTY0214">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean エラーコード    True:正常  False:異常</returns>
    ''' <remarks>指定されたＣＳＶファイルをコンボボックスにセットする
    ''' <para>作成情報：2016.01.20 y.morooka
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function GetCmbCsvData(ByRef dataEXTZ0214 As DataEXTZ0214)

        '開始ログ出力()
        commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        ' 変数宣言
        Dim cnt As New Integer

        Try
            'G請求先部署
            ' CSV読み込みクラスのインスタンス作成
            Using txtParser As New Microsoft.VisualBasic.FileIO.TextFieldParser(
                ConfigurationManager.AppSettings("csvSeikyusakiBusyoPath"), System.Text.Encoding.GetEncoding(932))

                txtParser.TextFieldType = FileIO.FieldType.Delimited
                txtParser.SetDelimiters(",")

                '読み込む行がなくなるまで繰り返し
                Dim strAryBuffer As String() = Nothing

                'DataTableオブジェクトを用意
                Dim seikyusakiBusyoTable As New DataTable()

                'DataTableに列を追加
                seikyusakiBusyoTable.Columns.Add("ID", GetType(String))
                seikyusakiBusyoTable.Columns.Add("NAME", GetType(String))

                While Not txtParser.EndOfData
                    cnt += 1
                    strAryBuffer = txtParser.ReadFields() '1行読み込み

                    If cnt <> 1 Then
                        If strAryBuffer(1) = dataEXTZ0214.PropStrAitesakiCD Then

                            '新しい行を作成
                            Dim row As DataRow = seikyusakiBusyoTable.NewRow()
                            '各列に値をセット
                            row("ID") = strAryBuffer(3)
                            row("NAME") = strAryBuffer(4)
                            'DataTableに行を追加
                            seikyusakiBusyoTable.Rows.Add(row)
                        End If
                    End If
                End While
                seikyusakiBusyoTable.AcceptChanges()
                '取得データをデータクラスへ保存
                dataEXTZ0214.PropTblSeikyusakiBusyo = seikyusakiBusyoTable
                '回数コンボボックス作成
                If commonLogic.SetCmbBox(dataEXTZ0214.PropTblSeikyusakiBusyo, dataEXTZ0214.PropCmbSeikyusakiBusyo) = False Then
                    Return False
                End If
                'コンボボックス先頭行のオプションの変更。
                If commonLogic.SetCmbBox(dataEXTZ0214.PropTblSeikyusakiBusyo, dataEXTZ0214.PropCmbSeikyusakiBusyo, False) = False Then
                    Return False
                End If

            End Using

            'G請求内容
            ' CSV読み込みクラスのインスタンス作成
            Using txtParser As New Microsoft.VisualBasic.FileIO.TextFieldParser(
               ConfigurationManager.AppSettings("csvSeikyuNaiyoPath"), System.Text.Encoding.GetEncoding(932))
                txtParser.TextFieldType = FileIO.FieldType.Delimited
                txtParser.SetDelimiters(",")

                '読み込む行がなくなるまで繰り返し
                Dim strAryBuffert As String() = Nothing

                'DataTableオブジェクトを用意)
                Dim SeikyuNaiyoTable As New DataTable()

                'DataTableに列を追加
                SeikyuNaiyoTable.Columns.Add("ID", GetType(String))
                SeikyuNaiyoTable.Columns.Add("NAME", GetType(String))
                cnt = 0
                While Not txtParser.EndOfData
                    cnt += 1
                    strAryBuffert = txtParser.ReadFields() '1行読み込み

                    If cnt <> 1 Then
                        If strAryBuffert(1) = dataEXTZ0214.PropStrAitesakiCD Then

                            '新しい行を作成
                            Dim row As DataRow = SeikyuNaiyoTable.NewRow()
                            '各列に値をセット
                            row("ID") = strAryBuffert(3)
                            row("NAME") = strAryBuffert(5)
                            'DataTableに行を追加
                            SeikyuNaiyoTable.Rows.Add(row)
                        End If
                    End If
                End While
                SeikyuNaiyoTable.AcceptChanges()
                '取得データをデータクラスへ保存
                dataEXTZ0214.PropTblSeikyuNaiyo = SeikyuNaiyoTable
                '回数コンボボックス作成
                If commonLogic.SetCmbBox(dataEXTZ0214.PropTblSeikyuNaiyo, dataEXTZ0214.PropCmbSeikyuNaiyo) = False Then
                    Return False
                End If
                'コンボボックス先頭行のオプションの変更。
                If commonLogic.SetCmbBox(dataEXTZ0214.PropTblSeikyuNaiyo, dataEXTZ0214.PropCmbSeikyuNaiyo, False) = False Then
                    Return False
                End If

            End Using

            '終了ログ出力
            commonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            Return True

        Catch ex As Exception
            commonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            puErrMsg = EXT_E001 & ex.Message
            Return False

        End Try

    End Function

End Class
