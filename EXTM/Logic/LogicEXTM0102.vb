Imports Common
Imports Common.CommonLogic
Imports Npgsql
Imports FarPoint.Win.Spread.CellType
Imports System.Text.RegularExpressions

Public Class LogicEXTM0102

    Private sqlEXTM0102 As New SqlEXTM0102
    Private dataEXTM0102 As New DataEXTM0102
    Private CommonLogic As New CommonLogic

    'コンボボックス
    Dim cmbKamoku As New FarPoint.Win.Spread.CellType.ComboBoxCellType()        '勘定科目
    'Dim cmbKarikamoku As New FarPoint.Win.Spread.CellType.ComboBoxCellType()    '借方勘定科目    ' 2016.04.28 DEL h.hagiwara レスポンス改善

    'コンボボックス作成用配列
    Dim aryNm As ArrayList = New ArrayList()
    Dim aryData As ArrayList = New ArrayList()

    ' 付帯設備Spread列位置
    Private Const M0102_FUTAI_COL_SETUBICD As Integer = 0
    Private Const M0102_FUTAI_COL_SETUBINM As Integer = 1
    Private Const M0102_FUTAI_COL_TANKA As Integer = 2
    Private Const M0102_FUTAI_COL_TANNI As Integer = 3
    Private Const M0102_FUTAI_COL_DEFFLG As Integer = 4
    Private Const M0102_FUTAI_COL_SORT As Integer = 5
    Private Const M0102_FUTAI_COL_STS As Integer = 6
    Private Const M0102_FUTAI_COL_ROW As Integer = 7
    Private Const M0102_FUTAI_COL_BUNRUICD As Integer = 8
    Private Const M0102_FUTAI_COL_UPKBN As Integer = 9

    ' 付帯分類Spread列位置       ' 2016.04.28 ADD h.hagiwara 
    Private Const M0102_BUNRUI_COL_BUNRUICD As Integer = 0
    Private Const M0102_BUNRUI_COL_BUNRUINM As Integer = 1
    Private Const M0102_BUNRUI_COL_SYUKEIKY As Integer = 2
    Private Const M0102_BUNRUI_COL_TAXKBN As Integer = 3
    Private Const M0102_BUNRUI_COL_KANJYO As Integer = 4
    Private Const M0102_BUNRUI_COL_SAIMOKU As Integer = 5
    Private Const M0102_BUNRUI_COL_UCHIWAKE As Integer = 6
    Private Const M0102_BUNRUI_COL_SYOSAI As Integer = 7
    Private Const M0102_BUNRUI_COL_SORT As Integer = 8
    Private Const M0102_BUNRUI_COL_DEFFLG As Integer = 9
    Private Const M0102_BUNRUI_COL_FRIYOFLG As Integer = 10
    ' 付帯分類Spread列位置       ' 2016.04.28 ADD h.hagiwara 

    ''' <summary>
    ''' 付帯設備マスタメンテ初期表示処理
    ''' <paramref name="dataEXTM0102">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean  取得合否　True:取得成功 Flase:取得失敗</returns>
    ''' <remarks>init 分類マスタより、データを取得
    ''' <para>作成情報：2015/08/13 yu.satoh
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function InitBunrui(ByRef dataEXTM0102 As DataEXTM0102) As Boolean

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)    'コネクション
        Dim Adapter As New NpgsqlDataAdapter        'アダプタ
        Dim Table As New DataTable()                'テーブル

        Try

            'コネクションを開く
            Cn.Open()

            '初期チェック実装
            If dataEXTM0102.PropInitFlg = False Then
                dataEXTM0102.PropTheaterBtn.Checked = True  'シアターにチェック
                dataEXTM0102.PropFinishedBtn.Checked = True '編集済みをチェック
            End If

            '付帯設備マスタメンテ用SQLの作成・設定
            If sqlEXTM0102.SetSelectIniData(Adapter, Cn, dataEXTM0102) = False Then
                Return False
            End If

            'SQLログ
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "システム管理情報", Nothing, Adapter.SelectCommand)

            'データを取得
            Adapter.Fill(Table)
            '取得データをデータクラスへ保存
            dataEXTM0102.PropDtFbunruiMst = Table

            ''対象期間を取得
            'If CmbFinishedFromTo(dataEXTM0102) = False Then
            '    Return False
            'End If

            ''SQLログ
            'CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "システム管理情報", Nothing, Adapter.SelectCommand)

            ''対象期間コンボボックス作成
            ''現在日付に対応するものがあれば、それを初期値に、そうでない場合は初期値を未設定で作成
            'If dataEXTM0102.PropDtFbunruiMst.Rows.Count <> 0 Then
            '    If CommonLogic.SetCmbBox(dataEXTM0102.PropDtFinishedFromTo, dataEXTM0102.PropFinishedFromTo, False, dataEXTM0102.PropDtFbunruiMst.Rows(0).Item(0), dataEXTM0102.PropDtFbunruiMst.Rows(0).Item(1)) = False Then
            '        Return False
            '    End If
            'Else
            '    If CommonLogic.SetCmbBox(dataEXTM0102.PropDtFinishedFromTo, dataEXTM0102.PropFinishedFromTo, True) = False Then
            '        Return False
            '    End If
            'End If

            '初期表示時処理
            If dataEXTM0102.PropInitCmbFlg = False Then

                '対象期間を取得
                If CmbFinishedFromTo(dataEXTM0102) = False Then
                    Return False
                End If

                'SQLログ
                CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "システム管理情報", Nothing, Adapter.SelectCommand)

                '対象期間コンボボックス作成
                '現在日付に対応するものがあれば、それを初期値に、そうでない場合は初期値を未設定で作成
                If dataEXTM0102.PropDtFbunruiMst.Rows.Count <> 0 Then
                    If CommonLogic.SetCmbBox(dataEXTM0102.PropDtFinishedFromTo, dataEXTM0102.PropFinishedFromTo, False, dataEXTM0102.PropDtFbunruiMst.Rows(0).Item(0), dataEXTM0102.PropDtFbunruiMst.Rows(0).Item(1)) = False Then
                        Return False
                    End If
                Else
                    If CommonLogic.SetCmbBox(dataEXTM0102.PropDtFinishedFromTo, dataEXTM0102.PropFinishedFromTo, True) = False Then
                        Return False
                    End If
                End If


                'SQLログ
                CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "システム管理情報", Nothing, Adapter.SelectCommand)

                ' 2016.02.12 DEL START↓ h.hagiwara
                ''分類表の勘定科目コード、借方科目コードを全てコンボボックスに指定
                'For i = 0 To dataEXTM0102.PropVwGroupingSheet.ActiveSheet.RowCount - 1
                '    'i4とi8はここで全てコンボボックスとして宣言
                '    dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 4).CellType = cmbKamoku
                '    dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 8).CellType = cmbKarikamoku
                'Next
                ' 2016.02.12 DEL END↑ h.hagiwara

                '勘定科目取得
                If CmbKamokuSet(dataEXTM0102) = False Then
                    Return False
                End If

                ' 2016.06.24 ADD START↓ h.hagiwara コンボリスト設定方法変更対応
                '細目取得
                If CmbSaimokuSet(dataEXTM0102) = False Then
                    Return False
                End If
                '内訳取得
                If CmbUchiwakeSet(dataEXTM0102) = False Then
                    Return False
                End If
                '詳細取得
                If CmbShosaiSet(dataEXTM0102) = False Then
                    Return False
                End If
                ' 2016.06.24 ADD END↑ h.hagiwara コンボリスト設定方法変更対応

                'SQLログ
                CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "システム管理情報", Nothing, Adapter.SelectCommand)

                '' 2016.02.12 DEL START↓ h.hagiwara
                ' ''勘定科目ボックス作成
                ''CmbKamokuCdCreate(dataEXTM0102)
                '' 2016.02.12 DEL END↑ h.hagiwara

                ' 2016.04.28 DEL START↓ h.hagiwara レスポンス改善
                ''借方勘定科目
                'If CmbKarikamokuSet(dataEXTM0102) = False Then
                '    Return False
                'End If

                ''SQLログ
                'CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "システム管理情報", Nothing, Adapter.SelectCommand)
                ' 2016.04.28 DEL END↑ h.hagiwara レスポンス改善

                ' 2016.02.12 DEL START↓ h.hagiwara
                ''借方勘定科目作成
                'CmbKarikamokuCreate(dataEXTM0102)
                ' 2016.02.12 DEL END↑ h.hagiwara

                '初期表示管理フラグを設定
                dataEXTM0102.PropInitCmbFlg = True
            End If

            '分類マスタを代入
            If SetBunruiData(dataEXTM0102) = False Then
                Return False
            End If

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            Return True
        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            puErrMsg = EXTM0102_E0000 & ex.Message
            Return False
        Finally
            Cn.Close()
            Cn.Dispose()
            Adapter.Dispose()
            Table.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' 付帯設備マスタメンテ初期表示処理
    ''' <paramref name="dataEXTM0102">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean  取得合否　True:取得成功 Flase:取得失敗</returns>
    ''' <remarks>init 分類マスタより、値を入れる
    ''' <para>作成情報：2015/08/13 yu.satoh
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function SetBunruiData(ByRef dataEXTM0102 As DataEXTM0102) As Boolean

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        'チェックボックス
        Dim chkBoxTax As New CheckBoxCellType   '税
        Dim chkBoxSts As New CheckBoxCellType   '無効フラグ

        '一度全てを空にし、分類コードを非活性に設定
        '現在日付に対応するものがあれば、それを初期値に、そうでない場合は初期値を未設定で作成
        If dataEXTM0102.PropDtFbunruiMst.Rows.Count <> 0 Then
            dataEXTM0102.PropYearFrom.Text = dataEXTM0102.PropDtFbunruiMst.Rows(0).Item(1)
            dataEXTM0102.PropMonthFrom.Text = dataEXTM0102.PropDtFbunruiMst.Rows(0).Item(2)
            dataEXTM0102.PropYearTo.Text = dataEXTM0102.PropDtFbunruiMst.Rows(0).Item(3)
            dataEXTM0102.PropMonthTo.Text = dataEXTM0102.PropDtFbunruiMst.Rows(0).Item(4)
            '更新用退避
            dataEXTM0102.PropFromYearUp = dataEXTM0102.PropDtFbunruiMst.Rows(0).Item(1)
            dataEXTM0102.PropFromMonthUp = dataEXTM0102.PropDtFbunruiMst.Rows(0).Item(2)
            dataEXTM0102.PropToYearUp = dataEXTM0102.PropDtFbunruiMst.Rows(0).Item(3)
            dataEXTM0102.PropToMonthUp = dataEXTM0102.PropDtFbunruiMst.Rows(0).Item(4)
            'ElseIf dataEXTM0102.PropFinishedFromTo.Text <> Nothing And dataEXTM0102.PropFinishedFromTo.Text.Length = 21 Then
            '    dataEXTM0102.PropYearFrom.Text = dataEXTM0102.PropFinishedFromTo.Text.ToString.Substring(0, 4)
            '    dataEXTM0102.PropMonthFrom.Text = dataEXTM0102.PropFinishedFromTo.Text.ToString.Substring(5, 2)
            '    dataEXTM0102.PropYearTo.Text = dataEXTM0102.PropFinishedFromTo.Text.ToString.Substring(11, 4)
            '    dataEXTM0102.PropMonthTo.Text = dataEXTM0102.PropFinishedFromTo.Text.ToString.Substring(16, 2)
            '    '更新用退避
            '    dataEXTM0102.PropFromYearUp = dataEXTM0102.PropFinishedFromTo.Text.ToString.Substring(0, 4)
            '    dataEXTM0102.PropFromMonthUp = dataEXTM0102.PropFinishedFromTo.Text.ToString.Substring(5, 2)
            '    dataEXTM0102.PropToYearUp = dataEXTM0102.PropFinishedFromTo.Text.ToString.Substring(11, 4)
            '    dataEXTM0102.PropToMonthUp = dataEXTM0102.PropFinishedFromTo.Text.ToString.Substring(16, 2)
        Else

            dataEXTM0102.PropYearFrom.Text = ""
            dataEXTM0102.PropMonthFrom.Text = ""
            dataEXTM0102.PropYearTo.Text = ""
            dataEXTM0102.PropMonthTo.Text = ""
            dataEXTM0102.PropFromYearUp = ""
            dataEXTM0102.PropFromMonthUp = ""
            dataEXTM0102.PropToYearUp = ""
            dataEXTM0102.PropToMonthUp = ""
        End If

        ' 2016.02.12 UPD START↓ h.hagiwara  
        'For i = 0 To dataEXTM0102.PropVwGroupingSheet.ActiveSheet.RowCount - 1
        '    For j = 0 To dataEXTM0102.PropVwGroupingSheet.ActiveSheet.ColumnCount - 1
        '        dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, j).Value = Nothing
        '    Next
        '    'dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 0).Locked = False
        '    dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 8).Locked = False
        'Next
        If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.RowCount > 0 Then
            dataEXTM0102.PropVwGroupingSheet.Sheets(0).Rows.Remove(0, dataEXTM0102.PropVwGroupingSheet.ActiveSheet.RowCount)
        End If
        ' 2016.02.12 UPD END↑  h.hagiwara  

        '分類表の表示行数を、sql取得行数＋５に設定
        'dataEXTM0102.PropVwGroupingSheet.ActiveSheet.RowCount = dataEXTM0102.PropDtFbunruiMst.Rows.Count + 5

        Try

            'もしデータがなかった場合、以下の処理を全て行わない
            If dataEXTM0102.PropDtFbunruiMst.Rows.Count <> 0 Then

                '対象期間コンボボックス設定、日付fromTo設定
                dataEXTM0102.PropFinishedFromTo.SelectedValue = dataEXTM0102.PropDtFbunruiMst.Rows(0).Item(0)
                dataEXTM0102.PropYearFrom.Text = dataEXTM0102.PropDtFbunruiMst.Rows(0).Item(1)
                dataEXTM0102.PropMonthFrom.Text = dataEXTM0102.PropDtFbunruiMst.Rows(0).Item(2)
                dataEXTM0102.PropYearTo.Text = dataEXTM0102.PropDtFbunruiMst.Rows(0).Item(3)
                dataEXTM0102.PropMonthTo.Text = dataEXTM0102.PropDtFbunruiMst.Rows(0).Item(4)

                '分類表の表示行数を、sql取得行数＋５に設定
                'dataEXTM0102.PropVwGroupingSheet.ActiveSheet.RowCount = dataEXTM0102.PropDtFbunruiMst.Rows.Count + 5

                '発行したSQLより、各初期値を代入
                For i = 0 To dataEXTM0102.PropDtFbunruiMst.Rows.Count - 1
                    dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Rows.Add(dataEXTM0102.PropVwGroupingSheet.ActiveSheet.RowCount, 1)
                    '分類コードを非活性にする。
                    dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, M0102_BUNRUI_COL_BUNRUICD).Locked = True
                    '分類コード
                    dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, M0102_BUNRUI_COL_BUNRUICD).Text = dataEXTM0102.PropDtFbunruiMst.Rows(i).Item(6)
                    '分類名
                    dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, M0102_BUNRUI_COL_BUNRUINM).Text = dataEXTM0102.PropDtFbunruiMst.Rows(i).Item(7)
                    '集計
                    dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, M0102_BUNRUI_COL_SYUKEIKY).Text = dataEXTM0102.PropDtFbunruiMst.Rows(i).Item(8)
                    '税
                    dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, M0102_BUNRUI_COL_TAXKBN).CellType = chkBoxTax
                    If dataEXTM0102.PropDtFbunruiMst.Rows(i).Item(9) = 1 Then
                        dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, M0102_BUNRUI_COL_TAXKBN).Value = True
                    End If

                    '勘定科目
                    CmbKamokuCdCreate(dataEXTM0102)                                                                                                           ' セレクトボックスの中身を作成  2016.02.12 ADD h.hagiwara
                    dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, M0102_BUNRUI_COL_KANJYO).CellType = cmbKamoku                                       ' 2016.02.12 ADD h.hagiwara
                    dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, M0102_BUNRUI_COL_KANJYO).Value = dataEXTM0102.PropDtFbunruiMst.Rows(i).Item(11)
                    dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, M0102_BUNRUI_COL_KANJYO).Locked = False                                             ' 2016.02.12 ADD h.hagiwara 稼働後№19対応

                    '勘定科目がある行の借方科目の前まで活性化
                    If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, M0102_BUNRUI_COL_KANJYO).Value <> Nothing Then
                        dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, M0102_BUNRUI_COL_SAIMOKU).Locked = False
                        dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, M0102_BUNRUI_COL_UCHIWAKE).Locked = False
                        dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, M0102_BUNRUI_COL_SYOSAI).Locked = False
                    End If

                    dataEXTM0102.PropStrSelKAnjo = dataEXTM0102.PropDtFbunruiMst.Rows(i).Item(11)                                 ' 2016.06.24 ADD h.hagiwara コンボリスト設定方法変更対応
                    '細目コンボボックスを作成
                    'If CmbSaimokuSet(dataEXTM0102, dataEXTM0102.PropDtFbunruiMst.Rows(i).Item(11)) = True Then                   ' 2016.06.24 DEL h.hagiwara コンボリスト設定方法変更対応

                    '宣言
                    Dim cmbSaimoku As New FarPoint.Win.Spread.CellType.ComboBoxCellType()

                    'セレクトボックスの中身を作成
                    CmbSaimokuCreate(dataEXTM0102, cmbSaimoku)

                    '細目の項目をセット
                    dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, M0102_BUNRUI_COL_SAIMOKU).CellType = cmbSaimoku
                    dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, M0102_BUNRUI_COL_SAIMOKU).Value = dataEXTM0102.PropDtFbunruiMst.Rows(i).Item(12)
                    'End If                                                                                                     ' 2016.06.24 DEL h.hagiwara コンボリスト設定方法変更対応


                    dataEXTM0102.PropStrSelSaimoku = dataEXTM0102.PropDtFbunruiMst.Rows(i).Item(12)                             ' 2016.06.24 ADD h.hagiwara コンボリスト設定方法変更対応
                    '内訳をセット
                    'If CmbUchiwakeSet(dataEXTM0102, dataEXTM0102.PropDtFbunruiMst.Rows(i).Item(11), _                          ' 2016.06.24 DEL h.hagiwara コンボリスト設定方法変更対応
                    '                  dataEXTM0102.PropDtFbunruiMst.Rows(i).Item(12)) = True Then                              ' 2016.06.24 DEL h.hagiwara コンボリスト設定方法変更対応                                   
                    'セルタイプと、コンボボックスの設定
                    Dim cmbUchi As New FarPoint.Win.Spread.CellType.ComboBoxCellType()

                    CmbUchiCreate(dataEXTM0102, cmbUchi)

                    '値をセット
                    dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, M0102_BUNRUI_COL_UCHIWAKE).CellType = cmbUchi
                    dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, M0102_BUNRUI_COL_UCHIWAKE).Value = dataEXTM0102.PropDtFbunruiMst.Rows(i).Item(13)
                    'End If                                                                                                    ' 2016.06.24 DEL h.hagiwara コンボリスト設定方法変更対応

                    dataEXTM0102.PropStrSelUchiwake = dataEXTM0102.PropDtFbunruiMst.Rows(i).Item(13)                           ' 2016.06.24 ADD h.hagiwara コンボリスト設定方法変更対応
                    '詳細
                    'If CmbShosaiSet(dataEXTM0102, dataEXTM0102.PropDtFbunruiMst.Rows(i).Item(11), _                           ' 2016.06.24 DEL h.hagiwara コンボリスト設定方法変更対応 
                    '      dataEXTM0102.PropDtFbunruiMst.Rows(i).Item(12), _                                                   ' 2016.06.24 DEL h.hagiwara コンボリスト設定方法変更対応
                    '      dataEXTM0102.PropDtFbunruiMst.Rows(i).Item(13)) = True Then                                         ' 2016.06.24 DEL h.hagiwara コンボリスト設定方法変更対応

                    'セルタイプと、コンボボックスの設定
                    Dim cmbShosai As New FarPoint.Win.Spread.CellType.ComboBoxCellType()

                    'コンボボックスの中身を作成
                    CmbSyosaiCreate(dataEXTM0102, cmbShosai)

                    '値をセット
                    dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, M0102_BUNRUI_COL_SYOSAI).CellType = cmbShosai
                    dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, M0102_BUNRUI_COL_SYOSAI).Value = dataEXTM0102.PropDtFbunruiMst.Rows(i).Item(14)
                    'End If                                                                                                  ' 2016.06.24 DEL h.hagiwara コンボリスト設定方法変更対応

                    ' 2016.04.28 DEL START↓ h.hagiwara レスポンス改善
                    '' 2016.02.12 ADD START↓ h.hagiwara
                    'CmbKarikamokuCreate(dataEXTM0102)
                    'dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 8).CellType = cmbKarikamoku
                    '' 2016.02.12 ADD END↑ h.hagiwara

                    ''取得データが空では無ければ実行
                    'If dataEXTM0102.PropDtFbunruiMst.Rows(i).Item(15).ToString IsNot "" Then
                    '    '借方科目コード
                    '    dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 8).Value = dataEXTM0102.PropDtFbunruiMst.Rows(i).Item(15)
                    'End If

                    ''借方科目コードがあれば、細目を活性化
                    'If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 8).Value IsNot Nothing Then
                    '    dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 9).Locked = False
                    'End If

                    ''借方科目コードがnothingならば行わない
                    'If dataEXTM0102.PropDtFbunruiMst.Rows(i).Item(15).ToString IsNot "" Then
                    '    '借方細目
                    '    If CmbKariSaimokuSet(dataEXTM0102, dataEXTM0102.PropDtFbunruiMst.Rows(i).Item(15)) = True Then

                    '        'セルタイプと、コンボボックスの設定
                    '        Dim cmbKarisaimoku As New FarPoint.Win.Spread.CellType.ComboBoxCellType()

                    '        CmbKarisaimokuCreate(dataEXTM0102, cmbKarisaimoku)

                    '        '値をセット
                    '        dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 9).CellType = cmbKarisaimoku
                    '        dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 9).Value = dataEXTM0102.PropDtFbunruiMst.Rows(i).Item(16)
                    '    End If
                    '    '借方細目があれば内訳活性化
                    '    If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 9).Value IsNot Nothing Then
                    '        dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 10).Locked = False
                    '    End If

                    '    '借方細目コードがあれば実行
                    '    If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 9).Value IsNot Nothing Then
                    '        '借方内訳
                    '        If CmbKariUchiSet(dataEXTM0102, dataEXTM0102.PropDtFbunruiMst.Rows(i).Item(15), _
                    '                          dataEXTM0102.PropDtFbunruiMst.Rows(i).Item(16)) = True Then
                    '            'セルタイプと、コンボボックスの設定
                    '            Dim cmbKariuchi As New FarPoint.Win.Spread.CellType.ComboBoxCellType()

                    '            CmbKariuchiCreate(dataEXTM0102, cmbKariuchi)

                    '            '値をセット
                    '            dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 10).CellType = cmbKariuchi
                    '            dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 10).Value = dataEXTM0102.PropDtFbunruiMst.Rows(i).Item(17)
                    '        End If
                    '        '借方内訳があれば詳細を活性化
                    '        If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 10).Value IsNot Nothing Then
                    '            dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 11).Locked = False
                    '        End If
                    '    End If

                    '    '借方内訳があれば実行
                    '    If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 10).Value IsNot Nothing Then
                    '        '借方詳細
                    '        If CmbKariShosaiSet(dataEXTM0102, dataEXTM0102.PropDtFbunruiMst.Rows(i).Item(15), _
                    '                          dataEXTM0102.PropDtFbunruiMst.Rows(i).Item(16), _
                    '                          dataEXTM0102.PropDtFbunruiMst.Rows(i).Item(17)) = True Then
                    '            'セルタイプと、コンボボックスの設定
                    '            Dim cmbKarisyosai As New FarPoint.Win.Spread.CellType.ComboBoxCellType()

                    '            'コンボボックスの中身を作成
                    '            CmbKarishosaiCreate(dataEXTM0102, cmbKarisyosai)
                    '            '値をセット
                    '            dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 11).CellType = cmbKarisyosai
                    '            dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 11).Value = dataEXTM0102.PropDtFbunruiMst.Rows(i).Item(18)
                    '        End If
                    '    End If
                    'End If
                    ' 2016.04.28 DEL END↑ h.hagiwara レスポンス改善

                    '並び順
                    dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, M0102_BUNRUI_COL_SORT).Text = dataEXTM0102.PropDtFbunruiMst.Rows(i).Item(10)

                    '無効フラグ
                    dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, M0102_BUNRUI_COL_DEFFLG).CellType = chkBoxSts
                    If dataEXTM0102.PropDtFbunruiMst.Rows(i).Item(19) = 1 Then
                        dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, M0102_BUNRUI_COL_DEFFLG).Value = True
                    Else
                        dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, M0102_BUNRUI_COL_DEFFLG).Value = False
                    End If
                    ' 2015.11.30 ADD STATR↓ h.hagiwara
                    '付帯利用料フラグ
                    dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, M0102_BUNRUI_COL_FRIYOFLG).CellType = chkBoxSts
                    If dataEXTM0102.PropDtFbunruiMst.Rows(i).Item(20) = 1 Then
                        dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, M0102_BUNRUI_COL_FRIYOFLG).Value = True
                    Else
                        dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, M0102_BUNRUI_COL_FRIYOFLG).Value = False
                    End If
                    ' 2015.11.30 ADD END↑ h.hagiwara
                Next
            End If

            ' 2016.02.12 ADD START↓ h.hagiwara
            For j = 1 To 5
                dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Rows.Add(dataEXTM0102.PropVwGroupingSheet.ActiveSheet.RowCount, 1)
                CmbKamokuCdCreate(dataEXTM0102)
                dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(dataEXTM0102.PropVwGroupingSheet.ActiveSheet.RowCount - 1, M0102_BUNRUI_COL_KANJYO).CellType = cmbKamoku
                ' 2016.04.28 DEL START↓ h.hagiwara レスポンス改善
                'CmbKarikamokuCreate(dataEXTM0102)
                'dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(dataEXTM0102.PropVwGroupingSheet.ActiveSheet.RowCount - 1, 8).CellType = cmbKarikamoku
                ' 2016.04.28 DEL END↑ h.hagiwara レスポンス改善
            Next
            ' 2016.02.12 ADD END↑ h.hagiwara

            '初期表示完了
            dataEXTM0102.PropInitFlg = True

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            Return True
        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            puErrMsg = EXTM0102_E0000 & ex.Message
            Return False
        End Try
    End Function


    ''' <summary>
    ''' 付帯設備マスタメンテ初期表示処理
    ''' <paramref name="dataEXTM0102">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean  取得合否　True:取得成功 Flase:取得失敗</returns>
    ''' <remarks>init 分類マスタより、データを取得
    ''' <para>作成情報：2015/08/13 yu.satoh
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function InitFutai(ByRef dataEXTM0102 As DataEXTM0102, ByRef row As Integer) As Boolean

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)    'コネクション
        Dim Adapter As New NpgsqlDataAdapter        'アダプタ
        Dim Table As New DataTable()                'テーブル

        Try

            'コネクションを開く
            Cn.Open()

            '付帯設備シート用SQLの作成・設定
            If sqlEXTM0102.SetSelectIniFutaiData(Adapter, Cn, dataEXTM0102, row) = False Then
                Return False
            End If

            'SQLログ
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "システム管理情報", Nothing, Adapter.SelectCommand)

            'データを取得
            Adapter.Fill(Table)
            '取得データをデータクラスへ保存
            dataEXTM0102.PropDtFutaiMst = Table

            ''取得したデータを付帯設備シートに代入
            'If SetFutaiData(dataEXTM0102, row) = False Then
            '    Return False
            'End If

            If InitFutaiBunruiSet(dataEXTM0102) = False Then
                Return False
            End If

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            Return True
        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            puErrMsg = EXTM0102_E0000 & ex.Message
            Return False
        Finally
            Cn.Close()
            Cn.Dispose()
            Adapter.Dispose()
            Table.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' 付帯設備マスタメンテ初期表示処理
    ''' <paramref name="dataEXTM0102">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean  取得合否　True:取得成功 Flase:取得失敗</returns>
    ''' <remarks>init 付帯設備マスタより、値を入れる
    ''' <para>作成情報：2015/08/13 yu.satoh
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function SetFutaiData(ByRef dataEXTM0102 As DataEXTM0102, Optional ByRef row As Integer = 0) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try

            'チェックボックス
            Dim chkBoxDef As New CheckBoxCellType   'デフォルトセット
            Dim chkBoxSts As New CheckBoxCellType   '無効フラグ
            Dim dtrow1 As DataRow
            Dim dtrow2 As DataRow

            '一度表示項目をクリアする
            For i = 0 To dataEXTM0102.PropVwFutaiSheet.ActiveSheet.RowCount - 1
                For j = 0 To dataEXTM0102.PropVwFutaiSheet.ActiveSheet.ColumnCount - 1
                    dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, j).Value = Nothing
                Next
                '付帯コードを活性化
                'dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, 0).Locked = False
            Next

            If dataEXTM0102.PropDtFutaiMstDsp Is Nothing Then
                dataEXTM0102.PropDtFutaiMstDsp = dataEXTM0102.PropDtFutaiMst.Clone
            Else
                dataEXTM0102.PropDtFutaiMstDsp.Clear()
            End If
            If dataEXTM0102.PropDtFutaiMstEscp Is Nothing Then
                dataEXTM0102.PropDtFutaiMstEscp = dataEXTM0102.PropDtFutaiMst.Clone
            Else
                dataEXTM0102.PropDtFutaiMstEscp.Clear()
            End If
            ' 格納領域を表示用・退避用に分けて格納する。
            For j = 0 To dataEXTM0102.PropDtFutaiMst.Rows.Count - 1
                If dataEXTM0102.PropDtFutaiMst.Rows(j).Item(M0102_FUTAI_COL_BUNRUICD).ToString = "" Then
                    If dataEXTM0102.PropDtFutaiMst.Rows(j).Item(M0102_FUTAI_COL_ROW) = row Then
                        dtrow1 = dataEXTM0102.PropDtFutaiMstDsp.NewRow
                        For k = 0 To 9
                            dtrow1(k) = dataEXTM0102.PropDtFutaiMst.Rows(j).Item(k)
                        Next
                        dataEXTM0102.PropDtFutaiMstDsp.Rows.Add(dtrow1)
                    Else
                        dtrow2 = dataEXTM0102.PropDtFutaiMstEscp.NewRow
                        For k = 0 To 9
                            dtrow2(k) = dataEXTM0102.PropDtFutaiMst.Rows(j).Item(k)
                        Next
                        dataEXTM0102.PropDtFutaiMstEscp.Rows.Add(dtrow2)
                    End If
                ElseIf dataEXTM0102.PropDtFutaiMst.Rows(j).Item(M0102_FUTAI_COL_BUNRUICD).ToString = dataEXTM0102.PropBunruiCd Then
                    dtrow1 = dataEXTM0102.PropDtFutaiMstDsp.NewRow
                    For k = 0 To 9
                        dtrow1(k) = dataEXTM0102.PropDtFutaiMst.Rows(j).Item(k)
                    Next
                    dataEXTM0102.PropDtFutaiMstDsp.Rows.Add(dtrow1)
                Else
                    dtrow2 = dataEXTM0102.PropDtFutaiMstEscp.NewRow
                    For k = 0 To 9
                        dtrow2(k) = dataEXTM0102.PropDtFutaiMst.Rows(j).Item(k)
                    Next
                    dataEXTM0102.PropDtFutaiMstEscp.Rows.Add(dtrow2)
                End If
                'If dataEXTM0102.PropDtFutaiMst.Rows(j).Item(M0102_FUTAI_COL_BUNRUICD) = dataEXTM0102.PropBunruiCd Or _
                '   dataEXTM0102.PropDtFutaiMst.Rows(j).Item(M0102_FUTAI_COL_ROW) = row Then
                '    dtrow1 = dataEXTM0102.PropDtFutaiMstDsp.NewRow
                '    For k = 0 To 9
                '        dtrow1(k) = dataEXTM0102.PropDtFutaiMst.Rows(j).Item(k)
                '    Next
                '    dataEXTM0102.PropDtFutaiMstDsp.Rows.Add(dtrow1)
                'Else
                '    dtrow2 = dataEXTM0102.PropDtFutaiMstEscp.NewRow
                '    For k = 0 To 9
                '        dtrow2(k) = dataEXTM0102.PropDtFutaiMst.Rows(j).Item(k)
                '    Next
                '    dataEXTM0102.PropDtFutaiMstEscp.Rows.Add(dtrow2)
                'End If
            Next
            '分類表の表示行数を、sql取得行数＋５に設定
            dataEXTM0102.PropVwFutaiSheet.ActiveSheet.RowCount = dataEXTM0102.PropDtFutaiMstDsp.Rows.Count + 5

            'もしも値がなければ、以下の処理を行わない
            If dataEXTM0102.PropDtFutaiMstDsp.Rows.Count <> Nothing Then

                'テーブルに値を保代入
                'For i = 0 To dataEXTM0102.PropDtFutaiMst.Rows.Count - 1
                '    '付帯コードを非活性に
                '    dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, 0).Locked = True
                '    '付帯コード
                '    dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, 0).Text = dataEXTM0102.PropDtFutaiMst.Rows(i).Item(2)
                '    '付帯設備名
                '    dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, 1).Text = dataEXTM0102.PropDtFutaiMst.Rows(i).Item(3)
                '    '単価
                '    dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, 2).Text = dataEXTM0102.PropDtFutaiMst.Rows(i).Item(4)
                '    '単位
                '    dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, 3).Text = dataEXTM0102.PropDtFutaiMst.Rows(i).Item(5)
                '    'デフォルトセット
                '    dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, 4).CellType = chkBoxDef
                '    If dataEXTM0102.PropDtFutaiMst.Rows(i).Item(7) = 1 Then
                '        dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, 4).Value = True
                '    End If
                '    '並び順
                '    dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, 5).Text = dataEXTM0102.PropDtFutaiMst.Rows(i).Item(8)
                '    '無効フラグ
                '    dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, 6).CellType = chkBoxSts
                '    If dataEXTM0102.PropDtFutaiMst.Rows(i).Item(6) = 1 Then
                '        dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, 6).Value = True
                '    End If
                'Next
                For i = 0 To dataEXTM0102.PropDtFutaiMstDsp.Rows.Count - 1
                    '付帯コードを非活性に
                    dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, M0102_FUTAI_COL_SETUBICD).Locked = True
                    '付帯コード
                    dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, M0102_FUTAI_COL_SETUBICD).Text = dataEXTM0102.PropDtFutaiMstDsp.Rows(i).Item(M0102_FUTAI_COL_SETUBICD).ToString
                    '付帯設備名
                    dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, M0102_FUTAI_COL_SETUBINM).Text = dataEXTM0102.PropDtFutaiMstDsp.Rows(i).Item(M0102_FUTAI_COL_SETUBINM)
                    '単価
                    dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, M0102_FUTAI_COL_TANKA).Text = dataEXTM0102.PropDtFutaiMstDsp.Rows(i).Item(M0102_FUTAI_COL_TANKA)
                    '単位
                    dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, M0102_FUTAI_COL_TANNI).Text = dataEXTM0102.PropDtFutaiMstDsp.Rows(i).Item(M0102_FUTAI_COL_TANNI)
                    'デフォルトセット
                    dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, M0102_FUTAI_COL_DEFFLG).CellType = chkBoxDef
                    If dataEXTM0102.PropDtFutaiMstDsp.Rows(i).Item(M0102_FUTAI_COL_DEFFLG) = "1" Then
                        dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, M0102_FUTAI_COL_DEFFLG).Value = True
                    End If
                    '並び順
                    dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, M0102_FUTAI_COL_SORT).Text = dataEXTM0102.PropDtFutaiMstDsp.Rows(i).Item(M0102_FUTAI_COL_SORT)
                    '無効フラグ
                    dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, M0102_FUTAI_COL_STS).CellType = chkBoxSts
                    If dataEXTM0102.PropDtFutaiMstDsp.Rows(i).Item(M0102_FUTAI_COL_STS) = "1" Then
                        dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, 6).Value = True
                    End If
                    '行番号
                    dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, M0102_FUTAI_COL_ROW).Text = dataEXTM0102.PropDtFutaiMstDsp.Rows(i).Item(M0102_FUTAI_COL_ROW)
                    '分類コード
                    dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, M0102_FUTAI_COL_BUNRUICD).Text = dataEXTM0102.PropDtFutaiMstDsp.Rows(i).Item(M0102_FUTAI_COL_BUNRUICD).ToString
                    '更新区分
                    dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, M0102_FUTAI_COL_UPKBN).Text = dataEXTM0102.PropDtFutaiMstDsp.Rows(i).Item(M0102_FUTAI_COL_UPKBN)
                Next
            End If

            For i = 0 To dataEXTM0102.PropVwFutaiSheet.ActiveSheet.RowCount - 1
                '行番号
                dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, M0102_FUTAI_COL_ROW).Text = row
            Next

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            Return True
        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            puErrMsg = EXTM0102_E0000 & ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' 付帯設備マスタメンテ、表示内容を元に新規登録ボタン押下時処理
    ''' </summary>
    ''' <remarks>既に登録された内容を元に、新規登録を行う。
    ''' 新規に料金を設定するがチェックされ、対象期間コンボボックスが活性化。
    ''' 期間From、期間Toがそれぞれクリアされる。
    ''' <para>作成情報：2015/08/10 yu.satoh
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Sub PushNewEntryBtn(ByRef dataEXTM0102 As DataEXTM0102)

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        'ボタンが押された状態にする
        dataEXTM0102.PropNewEntryBtnFlg = True

        '項目活性、クリア処理
        dataEXTM0102.PropNewBtn.Checked = True          '新規登録
        dataEXTM0102.PropFinishedFromTo.Enabled = False  '対処期間
        dataEXTM0102.PropNewEntryBtn.Enabled = False  '対処期間
        '期間fromTo全てを空にクリア
        dataEXTM0102.PropYearFrom.Text = ""
        dataEXTM0102.PropMonthFrom.Text = ""
        dataEXTM0102.PropYearTo.Text = ""
        dataEXTM0102.PropMonthTo.Text = ""

        ' 2016.04.27 ADD START↓ h.hagiwara 更新対象の項目値クリア(期間重複判定不具合対応)
        dataEXTM0102.PropFromYearUp = ""
        dataEXTM0102.PropFromMonthUp = ""
        dataEXTM0102.PropToYearUp = ""
        dataEXTM0102.PropToMonthUp = ""
        ' 2016.04.27 ADD START↓ h.hagiwara 更新対象の項目値クリア(期間重複判定不具合対応)

        ' 分類スプレッドの分類コードクリア
        For i = 0 To dataEXTM0102.PropVwGroupingSheet.ActiveSheet.RowCount - 1
            dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, M0102_BUNRUI_COL_BUNRUICD).Text = ""
        Next
        ' 付帯設備スプレッドの設備コードクリア
        For i = 0 To dataEXTM0102.PropVwFutaiSheet.ActiveSheet.RowCount - 1
            dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, M0102_FUTAI_COL_SETUBICD).Text = ""
            dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, M0102_FUTAI_COL_BUNRUICD).Text = ""
        Next
        ' 付帯設備データテーブルの設備コードクリア
        For i = 0 To dataEXTM0102.PropDtFutaiMst.Rows.Count - 1
            dataEXTM0102.PropDtFutaiMst.Rows(i).Item(M0102_FUTAI_COL_SETUBICD) = ""
            dataEXTM0102.PropDtFutaiMst.Rows(i).Item(M0102_FUTAI_COL_BUNRUICD) = ""
        Next
        For i = 0 To dataEXTM0102.PropDtFutaiMstDsp.Rows.Count - 1
            dataEXTM0102.PropDtFutaiMstDsp.Rows(i).Item(M0102_FUTAI_COL_SETUBICD) = ""
            dataEXTM0102.PropDtFutaiMstDsp.Rows(i).Item(M0102_FUTAI_COL_BUNRUICD) = ""
        Next
        For i = 0 To dataEXTM0102.PropDtFutaiMstEscp.Rows.Count - 1
            dataEXTM0102.PropDtFutaiMstEscp.Rows(i).Item(M0102_FUTAI_COL_SETUBICD) = ""
            dataEXTM0102.PropDtFutaiMstEscp.Rows(i).Item(M0102_FUTAI_COL_BUNRUICD) = ""
        Next

        '押された状態解除
        dataEXTM0102.PropNewEntryBtnFlg = False

        '終了ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

    End Sub

    ''' <summary>
    ''' 付帯設備マスタメンテ、新規に料金を設定ボタンチェック時
    ''' </summary>
    ''' <remarks>新規に料金を設定する
    ''' <para>作成情報：2015/08/13 yu.satoh
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Sub CheckNewBtn(ByRef dataEXTM0102 As DataEXTM0102)

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        If dataEXTM0102.PropNewBtn.Checked = True Then

            'コンボボックスを初期化し、活性化、新規に登録のボタンも活性化
            dataEXTM0102.PropFinishedFromTo.Text = ""
            dataEXTM0102.PropFinishedFromTo.SelectedValue = -1
            dataEXTM0102.PropFinishedFromTo.Enabled = False
            dataEXTM0102.PropNewEntryBtn.Enabled = False
            '期間をクリア
            dataEXTM0102.PropYearFrom.Text = ""
            dataEXTM0102.PropMonthFrom.Text = ""
            dataEXTM0102.PropYearTo.Text = ""
            dataEXTM0102.PropMonthTo.Text = ""
            '検索テーブルをクリア
            dataEXTM0102.PropDtFbunruiMst.Clear()
            dataEXTM0102.PropDtFutaiMst.Clear()

            For i = 0 To dataEXTM0102.PropVwGroupingSheet.ActiveSheet.RowCount - 1
                dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, M0102_BUNRUI_COL_SAIMOKU).Locked = True
                dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, M0102_BUNRUI_COL_UCHIWAKE).Locked = True
                dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, M0102_BUNRUI_COL_SYOSAI).Locked = True
                ' 2016.04.28 DEL START↓ h.hagiwara レスポンス改善
                'dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 9).Locked = True
                'dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 10).Locked = True
                'dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 11).Locked = True
                ' 2016.04.28 DEL END↑ h.hagiwara レスポンス改善
            Next
        End If

        '終了ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

    End Sub


    ''' <summary>
    ''' 付帯設備マスタメンテ、設定済みの料金を設定するラジオボタンチェック時処理
    ''' </summary>
    ''' <remarks>設定済みの料金を設定する
    ''' <para>作成情報：2015/08/13 yu.satoh
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Sub CheckFinishedBtn(ByRef dataEXTM0102 As DataEXTM0102)

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        If dataEXTM0102.PropFinishedBtn.Checked = True Then

            '活性化、新規に登録のボタンも活性化
            dataEXTM0102.PropFinishedFromTo.Enabled = True
            If dataEXTM0102.PropDtFinishedFromTo.Rows.Count > 0 Then
                dataEXTM0102.PropNewEntryBtn.Enabled = True
            Else
                dataEXTM0102.PropNewEntryBtn.Enabled = False
            End If

            '初期表示処理を同じことをする
            dataEXTM0102.PropInitFlg = False
            dataEXTM0102.PropInitCmbFlg = False                                                         ' 2016.02.12 ADD h.hagiwara

            '分類シート記述
            If InitBunrui(dataEXTM0102) = False Then
                MsgBox(puErrMsg)
                Exit Sub
            End If
            '付帯シート記述
            If InitFutai(dataEXTM0102, 0) = False Then
                MsgBox(puErrMsg)
                Exit Sub
            End If
            dataEXTM0102.PropBunruiCd = dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(0, M0102_BUNRUI_COL_BUNRUICD).Value
            '取得したデータを付帯設備シートに代入
            If SetFutaiData(dataEXTM0102) = False Then
                MsgBox(puErrMsg)
                Exit Sub
            End If
            '初期表示処理完了フラグを元に戻す
            dataEXTM0102.PropInitFlg = True
        End If

        '終了ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

    End Sub

    ''' <summary>
    ''' 編集済みコンボボックス取得処理
    ''' <paramref name="dataEXTM0102">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean  取得合否　True:取得成功 Flase:取得失敗</returns>
    ''' <remarks> 分類マスタより、データを取得
    ''' <para>作成情報：2015/08/13 yu.satoh
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function CmbFinishedFromTo(ByRef dataEXTM0102 As DataEXTM0102) As Boolean

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)    'コネクション
        Dim Adapter As New NpgsqlDataAdapter        'アダプタ
        Dim Table As New DataTable()                'テーブル

        Try

            'コネクションを開く
            Cn.Open()

            '対象期間コンボボックス作成
            If sqlEXTM0102.SetFinishedFromToData(Adapter, Cn, dataEXTM0102) = False Then
                Return False
            End If

            'SQLログ
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "システム管理情報", Nothing, Adapter.SelectCommand)

            'データを取得
            Adapter.Fill(Table)
            '取得データをデータクラスへ保存
            dataEXTM0102.PropDtFinishedFromTo = Table

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            Return True
        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            puErrMsg = EXTM0102_E0000 & ex.Message
            Return False
        Finally
            Cn.Close()
            Cn.Dispose()
            Adapter.Dispose()
            'Table.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' 付帯設備マスタメンテ初期表示科目マスタよりコードのコンボボックス生成処理
    ''' <paramref name="dataEXTM0102">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean  取得合否　True:取得成功 Flase:取得失敗</returns>
    ''' <remarks>科目マスタより、データを取得
    ''' <para>作成情報：2015/08/13 yu.satoh
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function CmbKamokuSet(ByRef dataEXTM0102 As DataEXTM0102) As Boolean

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)    'コネクション
        Dim Adapter As New NpgsqlDataAdapter        'アダプタ
        Dim Table As New DataTable()                'テーブル

        Try

            'コネクションを開く
            Cn.Open()

            '付帯設備マスタメンテ用SQLの作成・設定
            If sqlEXTM0102.CmbKanzyoCdSet(Adapter, Cn, dataEXTM0102) = False Then
                Return False
            End If

            'SQLログ
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "システム管理情報", Nothing, Adapter.SelectCommand)


            'データを取得
            Adapter.Fill(Table)
            '取得データをデータクラスへ保存
            dataEXTM0102.PropDtKanzyoCd = Table

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            Return True
        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            puErrMsg = EXTM0102_E0000 & ex.Message
            Return False
        Finally
            Cn.Close()
            Cn.Dispose()
            Adapter.Dispose()
            Table.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' 付帯設備マスタメンテ初期表示科目マスタより細目コードのコンボボックス生成処理
    ''' <paramref name="dataEXTM0102">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean  取得合否　True:取得成功 Flase:取得失敗</returns>
    ''' <remarks>科目マスタより、データを取得
    ''' <para>作成情報：2015/08/13 yu.satoh
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function CmbSaimokuSet(ByRef dataEXTM0102 As DataEXTM0102) As Boolean                                                      ' 2016.06.24 UPD h.hagiwara コンボリスト設定方法変更対応
        'Public Function CmbSaimokuSet(ByRef dataEXTM0102 As DataEXTM0102, ByRef kamokuCd As String) As Boolean                       ' 2016.06.24 UPD h.hagiwara コンボリスト設定方法変更対応

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)    'コネクション
        Dim Adapter As New NpgsqlDataAdapter        'アダプタ
        Dim Table As New DataTable()                'テーブル

        Try

            'コネクションを開く
            Cn.Open()

            '付帯設備マスタメンテ用SQLの作成・設定
            'If sqlEXTM0102.CmbSaimouCdSet(Adapter, Cn, dataEXTM0102, kamokuCd) = False Then                                        ' 2016.06.24 UPD h.hagiwara コンボリスト設定方法変更対応
            If sqlEXTM0102.CmbSaimouCdSet(Adapter, Cn, dataEXTM0102) = False Then                                                   ' 2016.06.24 UPD h.hagiwara コンボリスト設定方法変更対応
                Return False
            End If

            'SQLログ
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "システム管理情報", Nothing, Adapter.SelectCommand)

            'データを取得
            Adapter.Fill(Table)
            '取得データをデータクラスへ保存
            'dataEXTM0102.PropDtKamokuMst = Table                                                                                  ' 2016.06.24 UPD h.hagiwara コンボリスト設定方法変更対応
            dataEXTM0102.PropDtSaimoku = Table                                                                                     ' 2016.06.24 UPD h.hagiwara コンボリスト設定方法変更対応

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            Return True
        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            puErrMsg = EXTM0102_E0000 & ex.Message
            Return False
        Finally
            Cn.Close()
            Cn.Dispose()
            Adapter.Dispose()
            Table.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' 付帯設備マスタメンテ初期表示科目マスタより細目コードのコンボボックス生成処理
    ''' <paramref name="dataEXTM0102">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <remarks>科目マスタより、データを取得し、表示する
    ''' <para>作成情報：2015/08/13 yu.satoh
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Sub CmbSaimokuCreate(ByRef dataEXTM0102 As DataEXTM0102, ByRef cmbSaimokuSub As ComboBoxCellType)

        ''開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '宣言
        Dim cmbSaimoku As New FarPoint.Win.Spread.CellType.ComboBoxCellType()

        'リストを空にする
        aryNm.Clear()
        aryData.Clear()

        cmbSaimoku.Items = New String() {""}
        cmbSaimoku.ItemData = New String() {Nothing}
        ' 2015.11.25 ADD START↓ h.hagiwara
        aryNm.AddRange(cmbSaimoku.Items)
        aryData.AddRange(cmbSaimoku.ItemData)
        ' 2015.11.25 ADD END↑ h.hagiwara

        ' 2016.06.24 ADD START↓ h.hagiwara コンボリスト設定方法変更対応
        Dim dr_Set() As DataRow = dataEXTM0102.PropDtSaimoku.Select(dataEXTM0102.PropDtSaimoku.Columns(0).Caption & "='" & dataEXTM0102.PropStrSelKanjo & "'")
        ' 2016.06.24 ADD END↑ h.hagiwara コンボリスト設定方法変更対応

        'テーブルにデータを代入
        'For j = 0 To dataEXTM0102.PropDtKamokuMst.Rows.Count - 1                             ' 2016.06.24 UPD h.hagiwara コンボリスト設定方法変更対応
        For j = 0 To dr_Set.Length - 1                                                        ' 2016.06.24 UPD h.hagiwara コンボリスト設定方法変更対応

            ' 2015.11.25 DEL START↓ h.hagiwara
            'aryNm.AddRange(cmbSaimoku.Items)
            'aryData.AddRange(cmbSaimoku.ItemData)
            ' 2015.11.25 DEL END↑ h.hagiwara

            ' 2016.06.24 UPD START↓ h.hagiwara コンボリスト設定方法変更対応
            'aryNm.Add(dataEXTM0102.PropDtKamokuMst.Rows(j).Item(1).ToString)
            'aryData.Add(dataEXTM0102.PropDtKamokuMst.Rows(j).Item(0).ToString)
            aryNm.Add(dr_Set(j).Item(2).ToString)
            aryData.Add(dr_Set(j).Item(1).ToString)
            ' 2016.06.24 UPD END↑ h.hagiwara コンボリスト設定方法変更対応

        Next

        'コンボボックスに値を設定
        cmbSaimoku.Items = CType(aryNm.ToArray(GetType(String)), String())
        cmbSaimoku.ItemData = CType(aryData.ToArray(GetType(String)), String())

        'セルから取得／設定する値はItemDataとする
        cmbSaimoku.EditorValue = FarPoint.Win.Spread.CellType.EditorValue.ItemData

        cmbSaimokuSub = cmbSaimoku

        '終了ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

    End Sub

    ''' <summary>
    ''' 付帯設備マスタメンテ初期表示科目マスタより内訳コードのコンボボックス生成処理
    ''' <paramref name="dataEXTM0102">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <remarks>科目マスタより、データを取得し、表示する
    ''' <para>作成情報：2015/08/24 yu.satoh
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Sub CmbUchiCreate(ByRef dataEXTM0102 As DataEXTM0102, ByRef cmbUchiSub As ComboBoxCellType)

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Dim cmbUchi As New FarPoint.Win.Spread.CellType.ComboBoxCellType()

        'リストを空にする
        aryNm.Clear()
        aryData.Clear()

        cmbUchi.Items = New String() {""}
        cmbUchi.ItemData = New String() {Nothing}
        ' 2015.11.25 ADD START↓ h.hagiwara
        aryNm.AddRange(cmbUchi.Items)
        aryData.AddRange(cmbUchi.ItemData)
        ' 2015.11.25 ADD END↑ h.hagiwara

        ' 2016.06.24 ADD START↓ h.hagiwara コンボリスト設定方法変更対応
        Dim dr_Set() As DataRow = dataEXTM0102.PropDtUchiwake.Select(dataEXTM0102.PropDtUchiwake.Columns(0).Caption & "='" & dataEXTM0102.PropStrSelKanjo & "' AND " & dataEXTM0102.PropDtUchiwake.Columns(1).Caption & "='" & dataEXTM0102.PropStrSelSaimoku & "'")
        ' 2016.06.24 ADD END↑ h.hagiwara コンボリスト設定方法変更対応

        'データの設定
        'For j = 0 To dataEXTM0102.PropDtKamokuMst.Rows.Count - 1                            ' 2016.06.24 UPD h.hagiwara コンボリスト設定方法変更対応
        For j = 0 To dr_Set.Length - 1                                                       ' 2016.06.24 UPD h.hagiwara コンボリスト設定方法変更対応

            ' 2015.11.25 DEL START↓ h.hagiwara
            'aryNm.AddRange(cmbUchi.Items)
            'aryData.AddRange(cmbUchi.ItemData)
            ' 2015.11.25 DEL END↑ h.hagiwara

            ' 2016.06.24 UPD START↓ h.hagiwara コンボリスト設定方法変更対応
            'aryNm.Add(dataEXTM0102.PropDtKamokuMst.Rows(j).Item(1).ToString)
            'aryData.Add(dataEXTM0102.PropDtKamokuMst.Rows(j).Item(0).ToString)
            aryNm.Add(dr_Set(j).Item(3).ToString)
            aryData.Add(dr_Set(j).Item(2).ToString)
            ' 2016.06.24 UPD END↑ h.hagiwara コンボリスト設定方法変更対応

        Next

        '値を設定
        cmbUchi.Items = CType(aryNm.ToArray(GetType(String)), String())
        cmbUchi.ItemData = CType(aryData.ToArray(GetType(String)), String())
        'セルから取得／設定する値はItemDataとする
        cmbUchi.EditorValue = FarPoint.Win.Spread.CellType.EditorValue.ItemData

        cmbUchiSub = cmbUchi

        '終了ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

    End Sub

    ''' <summary>
    ''' 付帯設備マスタメンテ初期表示科目マスタより詳細コードのコンボボックス生成処理
    ''' <paramref name="dataEXTM0102">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <remarks>科目マスタより、データを取得し、表示する
    ''' <para>作成情報：2015/08/24 yu.satoh
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Sub CmbSyosaiCreate(ByRef dataEXTM0102 As DataEXTM0102, ByRef CmbShosaiSub As ComboBoxCellType)

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Dim cmbShosai As New FarPoint.Win.Spread.CellType.ComboBoxCellType()        '詳細

        'リストを空にする
        aryNm.Clear()
        aryData.Clear()

        cmbShosai.Items = New String() {""}
        cmbShosai.ItemData = New String() {Nothing}
        ' 2015.11.25 ADD START↓ h.hagiwara
        aryNm.AddRange(cmbShosai.Items)
        aryData.AddRange(cmbShosai.ItemData)
        ' 2015.11.25 ADD END↑ h.hagiwara

        ' 2016.06.24 ADD START↓ h.hagiwara コンボリスト設定方法変更対応
        Dim dr_Set() As DataRow = dataEXTM0102.PropDtShosai.Select(dataEXTM0102.PropDtShosai.Columns(0).Caption & "='" & dataEXTM0102.PropStrSelKanjo & "' AND " & dataEXTM0102.PropDtShosai.Columns(1).Caption & "='" & dataEXTM0102.PropStrSelSaimoku & "' AND " & dataEXTM0102.PropDtShosai.Columns(2).Caption & "='" & dataEXTM0102.PropStrSelUchiwake & "'")
        ' 2016.06.24 ADD END↑ h.hagiwara コンボリスト設定方法変更対応

        'データの設定
        'For j = 0 To dataEXTM0102.PropDtKamokuMst.Rows.Count - 1                            ' 2016.06.24 UPD h.hagiwara コンボリスト設定方法変更対応
        For j = 0 To dr_Set.Length - 1                                                       ' 2016.06.24 UPD h.hagiwara コンボリスト設定方法変更対応

            ' 2015.11.25 DEL START↓ h.hagiwara
            'aryNm.AddRange(cmbShosai.Items)
            'aryData.AddRange(cmbShosai.ItemData)
            ' 2015.11.25 DEL END↑ h.hagiwara
            ' 2016.06.24 UPD START↓ h.hagiwara コンボリスト設定方法変更対応
            'aryNm.Add(dataEXTM0102.PropDtKamokuMst.Rows(j).Item(1).ToString)
            'aryData.Add(dataEXTM0102.PropDtKamokuMst.Rows(j).Item(0).ToString)
            aryNm.Add(dr_Set(j).Item(4).ToString)
            aryData.Add(dr_Set(j).Item(3).ToString)
            ' 2016.06.24 UPD END↑ h.hagiwara コンボリスト設定方法変更対応

        Next

        'コンボボックスに値を設定
        cmbShosai.Items = CType(aryNm.ToArray(GetType(String)), String())
        cmbShosai.ItemData = CType(aryData.ToArray(GetType(String)), String())

        'セルから取得／設定する値はItemDataとする
        cmbShosai.EditorValue = FarPoint.Win.Spread.CellType.EditorValue.ItemData

        CmbShosaiSub = cmbShosai

        '終了ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

    End Sub

    ' 2016.04.28 DEL START↓ h.hagiwara レスポンス改善
    ' ''' <summary>
    ' ''' 付帯設備マスタメンテ初期表示科目マスタより借方細目コードのコンボボックス生成処理
    ' ''' <paramref name="dataEXTM0102">[IN/OUT]データクラス</paramref>
    ' ''' </summary>
    ' ''' <remarks>科目マスタより、データを取得し、表示する
    ' ''' <para>作成情報：2015/08/24 yu.satoh
    ' ''' <p>改訂情報 : </p>
    ' ''' </para></remarks>
    'Public Sub CmbKarisaimokuCreate(ByRef dataEXTM0102 As DataEXTM0102, ByRef cmbKarisaimokuSub As ComboBoxCellType)

    '    '開始ログ出力()
    '    CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

    '    Dim cmbKarisaimoku As New FarPoint.Win.Spread.CellType.ComboBoxCellType()   '借方細目

    '    'リストを空にする
    '    aryNm.Clear()
    '    aryData.Clear()

    '    cmbKarisaimoku.Items = New String() {""}
    '    cmbKarisaimoku.ItemData = New String() {Nothing}
    '    ' 2015.11.25 ADD START↓ h.hagiwara
    '    aryNm.AddRange(cmbKarisaimoku.Items)
    '    aryData.AddRange(cmbKarisaimoku.ItemData)
    '    ' 2015.11.25 ADD END↑ h.hagiwara

    '    For j = 0 To dataEXTM0102.PropDtKamokuMst.Rows.Count - 1

    '        ' 2015.11.25 DEL START↓ h.hagiwara
    '        'aryNm.AddRange(cmbKarisaimoku.Items)
    '        'aryData.AddRange(cmbKarisaimoku.ItemData)
    '        ' 2015.11.25 ADD END↑ h.hagiwara
    '        aryNm.Add(dataEXTM0102.PropDtKamokuMst.Rows(j).Item(1).ToString)
    '        aryData.Add(dataEXTM0102.PropDtKamokuMst.Rows(j).Item(0).ToString)

    '    Next

    '    'コンボボックスにセット
    '    cmbKarisaimoku.Items = CType(aryNm.ToArray(GetType(String)), String())
    '    cmbKarisaimoku.ItemData = CType(aryData.ToArray(GetType(String)), String())

    '    'セルから取得／設定する値はItemDataとする
    '    cmbKarisaimoku.EditorValue = FarPoint.Win.Spread.CellType.EditorValue.ItemData

    '    cmbKarisaimokuSub = cmbKarisaimoku

    '    '終了ログ出力
    '    CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

    'End Sub

    ' ''' <summary>
    ' ''' 付帯設備マスタメンテ初期表示科目マスタより借方内訳コードのコンボボックス生成処理
    ' ''' <paramref name="dataEXTM0102">[IN/OUT]データクラス</paramref>
    ' ''' </summary>
    ' ''' <remarks>科目マスタより、データを取得し、表示する
    ' ''' <para>作成情報：2015/08/24 yu.satoh
    ' ''' <p>改訂情報 : </p>
    ' ''' </para></remarks>
    'Public Sub CmbKariuchiCreate(ByRef dataEXTM0102 As DataEXTM0102, ByRef cmbKariuchiSub As ComboBoxCellType)

    '    '開始ログ出力()
    '    CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

    '    Dim cmbKariuchi As New FarPoint.Win.Spread.CellType.ComboBoxCellType()   '借方細目

    '    'リストを空にする
    '    aryNm.Clear()
    '    aryData.Clear()

    '    cmbKariuchi.Items = New String() {""}
    '    cmbKariuchi.ItemData = New String() {Nothing}
    '    ' 2015.11.25 ADD START↓ h.hagiwara
    '    aryNm.AddRange(cmbKariuchi.Items)
    '    aryData.AddRange(cmbKariuchi.ItemData)
    '    ' 2015.11.25 ADD END↑ h.hagiwara

    '    For j = 0 To dataEXTM0102.PropDtKamokuMst.Rows.Count - 1

    '        ' 2015.11.25 DEL START↓ h.hagiwara
    '        'aryNm.AddRange(cmbKariuchi.Items)
    '        'aryData.AddRange(cmbKariuchi.ItemData)
    '        ' 2015.11.25 ADD END↑ h.hagiwara
    '        aryNm.Add(dataEXTM0102.PropDtKamokuMst.Rows(j).Item(1).ToString)
    '        aryData.Add(dataEXTM0102.PropDtKamokuMst.Rows(j).Item(0).ToString)

    '    Next

    '    'コンボボックスにセット
    '    cmbKariuchi.Items = CType(aryNm.ToArray(GetType(String)), String())
    '    cmbKariuchi.ItemData = CType(aryData.ToArray(GetType(String)), String())

    '    'セルから取得／設定する値はItemDataとする
    '    cmbKariuchi.EditorValue = FarPoint.Win.Spread.CellType.EditorValue.ItemData

    '    cmbKariuchiSub = cmbKariuchi

    '    '終了ログ出力
    '    CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

    'End Sub

    ' ''' <summary>
    ' ''' 付帯設備マスタメンテ初期表示科目マスタより借方詳細コードのコンボボックス生成処理
    ' ''' <paramref name="dataEXTM0102">[IN/OUT]データクラス</paramref>
    ' ''' </summary>
    ' ''' <remarks>科目マスタより、データを取得し、表示する
    ' ''' <para>作成情報：2015/08/24 yu.satoh
    ' ''' <p>改訂情報 : </p>
    ' ''' </para></remarks>
    'Public Sub CmbKarishosaiCreate(ByRef dataEXTM0102 As DataEXTM0102, ByRef cmbKarisyosaiSub As ComboBoxCellType)

    '    '開始ログ出力()
    '    CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

    '    Dim cmbKarisyosai As New FarPoint.Win.Spread.CellType.ComboBoxCellType()    '借方詳細

    '    'リストを空にする
    '    aryNm.Clear()
    '    aryData.Clear()

    '    cmbKarisyosai.Items = New String() {""}
    '    cmbKarisyosai.ItemData = New String() {Nothing}
    '    ' 2015.11.25 ADD START↓ h.hagiwara
    '    aryNm.AddRange(cmbKarisyosai.Items)
    '    aryData.AddRange(cmbKarisyosai.ItemData)
    '    ' 2015.11.25 ADD END↑ h.hagiwara

    '    For j = 0 To dataEXTM0102.PropDtKamokuMst.Rows.Count - 1

    '        ' 2015.11.25 DEL START↓ h.hagiwara
    '        'aryNm.AddRange(cmbKarisyosai.Items)
    '        'aryData.AddRange(cmbKarisyosai.ItemData)
    '        ' 2015.11.25 ADD END↑ h.hagiwara
    '        aryNm.Add(dataEXTM0102.PropDtKamokuMst.Rows(j).Item(1).ToString)
    '        aryData.Add(dataEXTM0102.PropDtKamokuMst.Rows(j).Item(0).ToString)

    '    Next

    '    'コンボボックスにセット
    '    cmbKarisyosai.Items = CType(aryNm.ToArray(GetType(String)), String())
    '    cmbKarisyosai.ItemData = CType(aryData.ToArray(GetType(String)), String())

    '    'セルから取得／設定する値はItemDataとする
    '    cmbKarisyosai.EditorValue = FarPoint.Win.Spread.CellType.EditorValue.ItemData

    '    cmbKarisyosaiSub = cmbKarisyosai

    '    '終了ログ出力
    '    CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

    'End Sub
    ' 2016.04.28 DEL END↑ h.hagiwara レスポンス改善

    ''' <summary>
    ''' 付帯設備マスタメンテ初期表示科目マスタより内訳コードのコンボボックス生成処理
    ''' <paramref name="dataEXTM0102">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean  取得合否　True:取得成功 Flase:取得失敗</returns>
    ''' <remarks>科目マスタより、データを取得
    ''' <para>作成情報：2015/08/13 yu.satoh
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function CmbUchiwakeSet(ByRef dataEXTM0102 As DataEXTM0102) As Boolean                                                                                ' 2016.06.24 UPD h.hagiwara コンボリスト設定方法変更対応
        'Public Function CmbUchiwakeSet(ByRef dataEXTM0102 As DataEXTM0102, ByRef kamokuCd As String, ByRef saimokuCd As String) As Boolean                      ' 2016.06.24 UPD h.hagiwara コンボリスト設定方法変更対応

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)    'コネクション
        Dim Adapter As New NpgsqlDataAdapter        'アダプタ
        Dim Table As New DataTable()                'テーブル

        Try

            'コネクションを開く
            Cn.Open()

            '付帯設備マスタメンテ用SQLの作成・設定
            'If sqlEXTM0102.CmbUchiwakeCdSet(Adapter, Cn, dataEXTM0102, kamokuCd, saimokuCd) = False Then                                  ' 2016.06.24 UPD h.hagiwara コンボリスト設定方法変更対応
            If sqlEXTM0102.CmbUchiwakeCdSet(Adapter, Cn, dataEXTM0102) = False Then                                                        ' 2016.06.24 UPD h.hagiwara コンボリスト設定方法変更対応
                Return False
            End If

            'SQLログ
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "システム管理情報", Nothing, Adapter.SelectCommand)

            'データを取得
            Adapter.Fill(Table)
            '取得データをデータクラスへ保存
            'dataEXTM0102.PropDtKamokuMst = Table                                                                                         ' 2016.06.24 UPD h.hagiwara コンボリスト設定方法変更対応
            dataEXTM0102.PropDtUchiwake = Table                                                                                           ' 2016.06.24 UPD h.hagiwara コンボリスト設定方法変更対応

            ' MsgBox(Table.Rows(0).Item(0))

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            Return True
        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            puErrMsg = EXTM0102_E0000 & ex.Message
            Return False
        Finally
            Cn.Close()
            Cn.Dispose()
            Adapter.Dispose()
            Table.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' 付帯設備マスタメンテ初期表示科目マスタより詳細コードのコンボボックス生成処理
    ''' <paramref name="dataEXTM0102">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean  取得合否　True:取得成功 Flase:取得失敗</returns>
    ''' <remarks>科目マスタより、データを取得
    ''' <para>作成情報：2015/08/13 yu.satoh
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function CmbShosaiSet(ByRef dataEXTM0102 As DataEXTM0102) As Boolean                                                                                               ' 2016.06.24 UPD h.hagiwara コンボリスト設定方法変更対応
        'Public Function CmbShosaiSet(ByRef dataEXTM0102 As DataEXTM0102, ByRef kamokuCd As String, ByRef saimokuCd As String, ByRef uchiCd As String) As Boolean             ' 2016.06.24 UPD h.hagiwara コンボリスト設定方法変更対応

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)    'コネクション
        Dim Adapter As New NpgsqlDataAdapter        'アダプタ
        Dim Table As New DataTable()                'テーブル

        Try

            'コネクションを開く
            Cn.Open()

            '付帯設備マスタメンテ用SQLの作成・設定
            'If sqlEXTM0102.CmbShosaiCdSet(Adapter, Cn, dataEXTM0102, kamokuCd, saimokuCd, uchiCd) = False Then                           ' 2016.06.24 UPD h.hagiwara コンボリスト設定方法変更対応
            If sqlEXTM0102.CmbShosaiCdSet(Adapter, Cn, dataEXTM0102) = False Then                                                         ' 2016.06.24 UPD h.hagiwara コンボリスト設定方法変更対応
                Return False
            End If

            'SQLログ
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "システム管理情報", Nothing, Adapter.SelectCommand)

            'データを取得
            Adapter.Fill(Table)
            '取得データをデータクラスへ保存
            'dataEXTM0102.PropDtKamokuMst = Table                                                                                        ' 2016.06.24 UPD h.hagiwara コンボリスト設定方法変更対応
            dataEXTM0102.PropDtShosai = Table                                                                                            ' 2016.06.24 UPD h.hagiwara コンボリスト設定方法変更対応

            ' 2016.04.28 DEL START↓ h.hagiwara レスポンス改善
            ''コンボボックス作成
            'CmbKarikamokuSet(dataEXTM0102)
            ' 2016.04.28 DEL END↑ h.hagiwara レスポンス改善

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            Return True
        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            puErrMsg = EXTM0102_E0000 & ex.Message
            Return False
        Finally
            Cn.Close()
            Cn.Dispose()
            Adapter.Dispose()
            Table.Dispose()
        End Try

    End Function

    ' 2016.04.28 DEL START↓ h.hagiwara レスポンス改善
    ' ''' <summary>
    ' ''' 付帯設備マスタメンテ初期表示科目マスタより借方科目コードのコンボボックス生成処理
    ' ''' <paramref name="dataEXTM0102">[IN/OUT]データクラス</paramref>
    ' ''' </summary>
    ' ''' <returns>boolean  取得合否　True:取得成功 Flase:取得失敗</returns>
    ' ''' <remarks>科目マスタより、データを取得
    ' ''' <para>作成情報：2015/08/13 yu.satoh
    ' ''' <p>改訂情報 : </p>
    ' ''' </para></remarks>
    'Public Function CmbKarikamokuSet(ByRef dataEXTM0102 As DataEXTM0102) As Boolean

    '    '開始ログ出力()
    '    CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

    '    '変数宣言
    '    Dim Cn As New NpgsqlConnection(DbString)    'コネクション
    '    Dim Adapter As New NpgsqlDataAdapter        'アダプタ
    '    Dim Table As New DataTable()                'テーブル

    '    Try

    '        'コネクションを開く
    '        Cn.Open()

    '        '付帯設備マスタメンテ用SQLの作成・設定
    '        If sqlEXTM0102.CmbKariKamokuSet(Adapter, Cn, dataEXTM0102) = False Then
    '            Return False
    '        End If

    '        'SQLログ
    '        CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "システム管理情報", Nothing, Adapter.SelectCommand)


    '        'データを取得
    '        Adapter.Fill(Table)
    '        '取得データをデータクラスへ保存
    '        dataEXTM0102.PropDtKariKanzyoCd = Table

    '        '終了ログ出力
    '        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

    '        Return True
    '    Catch ex As Exception
    '        '例外発生
    '        CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
    '        puErrMsg = EXTM0102_E0000 & ex.Message
    '        Return False
    '    Finally
    '        Cn.Close()
    '        Cn.Dispose()
    '        Adapter.Dispose()
    '        Table.Dispose()
    '    End Try

    'End Function

    ' ''' <summary>
    ' ''' 付帯設備マスタメンテ初期表示科目マスタより借方科目コードのコンボボックス生成処理
    ' ''' <paramref name="dataEXTM0102">[IN/OUT]データクラス</paramref>
    ' ''' </summary>
    ' ''' <remarks>科目マスタより、データを取得しコンボボックス作成
    ' ''' <para>作成情報：2015/08/13 yu.satoh
    ' ''' <p>改訂情報 : </p>
    ' ''' </para></remarks>
    'Public Sub CmbKarikamokuCreate(ByRef dataEXTM0102 As DataEXTM0102)

    '    '開始ログ出力()
    '    CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

    '    cmbKarikamoku.Items = New String() {""}
    '    cmbKarikamoku.ItemData = New String() {Nothing}

    '    ' 2015.11.25 ADD START↓ h.hagiwara
    '    Dim aryNm As ArrayList = New ArrayList()
    '    Dim aryData As ArrayList = New ArrayList()
    '    aryNm.AddRange(cmbKarikamoku.Items)
    '    aryData.AddRange(cmbKarikamoku.ItemData)
    '    ' 2015.11.25 ADD END↑ h.hagiwara

    '    For i = 0 To dataEXTM0102.PropDtKariKanzyoCd.Rows.Count - 1
    '        ' 2015.11.25 DEL START↓ h.hagiwara
    '        'Dim aryNm As ArrayList = New ArrayList()
    '        'Dim aryData As ArrayList = New ArrayList()

    '        'aryNm.AddRange(cmbKarikamoku.Items)
    '        'aryData.AddRange(cmbKarikamoku.ItemData)
    '        ' 2015.11.25 DEL  END↑ h.hagiwara
    '        aryNm.Add(dataEXTM0102.PropDtKariKanzyoCd.Rows(i).Item(1).ToString)
    '        aryData.Add(dataEXTM0102.PropDtKariKanzyoCd.Rows(i).Item(0).ToString)

    '        ' 2015.11.25 DEL START↓ h.hagiwara
    '        'cmbKarikamoku.Items = CType(aryNm.ToArray(GetType(String)), String())
    '        'cmbKarikamoku.ItemData = CType(aryData.ToArray(GetType(String)), String())

    '        ''セルから取得／設定する値はItemDataとする
    '        'cmbKarikamoku.EditorValue = FarPoint.Win.Spread.CellType.EditorValue.ItemData
    '        ' 2015.11.25 DEL  END↑ h.hagiwara
    '    Next
    '    ' 2015.11.25 ADD START↓ h.hagiwara
    '    cmbKarikamoku.Items = CType(aryNm.ToArray(GetType(String)), String())
    '    cmbKarikamoku.ItemData = CType(aryData.ToArray(GetType(String)), String())
    '    'セルから取得／設定する値はItemDataとする
    '    cmbKarikamoku.EditorValue = FarPoint.Win.Spread.CellType.EditorValue.ItemData
    '    ' 2015.11.25 ADD END↑ h.hagiwara

    '    '終了ログ出力()
    '    CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

    'End Sub

    ' ''' <summary>
    ' ''' 付帯設備マスタメンテ初期表示科目マスタより借方細目コードのコンボボックス生成処理
    ' ''' <paramref name="dataEXTM0102">[IN/OUT]データクラス</paramref>
    ' ''' </summary>
    ' ''' <returns>boolean  取得合否　True:取得成功 Flase:取得失敗</returns>
    ' ''' <remarks>科目マスタより、データを取得
    ' ''' <para>作成情報：2015/08/13 yu.satoh
    ' ''' <p>改訂情報 : </p>
    ' ''' </para></remarks>
    'Public Function CmbKariSaimokuSet(ByRef dataEXTM0102 As DataEXTM0102, ByRef karikamokuCd As String) As Boolean

    '    '開始ログ出力()
    '    CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

    '    '変数宣言
    '    Dim Cn As New NpgsqlConnection(DbString)    'コネクション
    '    Dim Adapter As New NpgsqlDataAdapter        'アダプタ
    '    Dim Table As New DataTable()                'テーブル

    '    Try

    '        'コネクションを開く
    '        Cn.Open()

    '        '付帯設備マスタメンテ用SQLの作成・設定
    '        If sqlEXTM0102.CmbKariSaimoku(Adapter, Cn, dataEXTM0102, karikamokuCd) = False Then
    '            Return False
    '        End If

    '        'SQLログ
    '        CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "システム管理情報", Nothing, Adapter.SelectCommand)

    '        'データを取得
    '        Adapter.Fill(Table)
    '        '取得データをデータクラスへ保存
    '        dataEXTM0102.PropDtKamokuMst = Table

    '        '終了ログ出力
    '        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

    '        Return True
    '    Catch ex As Exception
    '        '例外発生
    '        CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
    '        puErrMsg = EXTM0102_E0000 & ex.Message
    '        Return False
    '    Finally
    '        Cn.Close()
    '        Cn.Dispose()
    '        Adapter.Dispose()
    '        Table.Dispose()
    '    End Try

    'End Function

    ' ''' <summary>
    ' ''' 付帯設備マスタメンテ初期表示科目マスタより借方内訳コードのコンボボックス生成処理
    ' ''' <paramref name="dataEXTM0102">[IN/OUT]データクラス</paramref>
    ' ''' </summary>
    ' ''' <returns>boolean  取得合否　True:取得成功 Flase:取得失敗</returns>
    ' ''' <remarks>科目マスタより、データを取得
    ' ''' <para>作成情報：2015/08/13 yu.satoh
    ' ''' <p>改訂情報 : </p>
    ' ''' </para></remarks>
    'Public Function CmbKariUchiSet(ByRef dataEXTM0102 As DataEXTM0102, ByRef karikamokuCd As String, ByRef kariSaimokuCd As String) As Boolean

    '    '開始ログ出力()
    '    CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

    '    '変数宣言
    '    Dim Cn As New NpgsqlConnection(DbString)    'コネクション
    '    Dim Adapter As New NpgsqlDataAdapter        'アダプタ
    '    Dim Table As New DataTable()                'テーブル

    '    Try

    '        'コネクションを開く
    '        Cn.Open()

    '        '付帯設備マスタメンテ用SQLの作成・設定
    '        If sqlEXTM0102.CmbKariUchi(Adapter, Cn, dataEXTM0102, karikamokuCd, kariSaimokuCd) = False Then
    '            Return False
    '        End If

    '        'SQLログ
    '        CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "システム管理情報", Nothing, Adapter.SelectCommand)


    '        'データを取得
    '        Adapter.Fill(Table)
    '        '取得データをデータクラスへ保存
    '        dataEXTM0102.PropDtKamokuMst = Table

    '        '終了ログ出力
    '        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

    '        Return True
    '    Catch ex As Exception
    '        '例外発生
    '        CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
    '        puErrMsg = EXTM0102_E0000 & ex.Message
    '        Return False
    '    Finally
    '        Cn.Close()
    '        Cn.Dispose()
    '        Adapter.Dispose()
    '        Table.Dispose()
    '    End Try

    'End Function

    ' ''' <summary>
    ' ''' 付帯設備マスタメンテ初期表示科目マスタより借方詳細コードのコンボボックス生成処理
    ' ''' <paramref name="dataEXTM0102">[IN/OUT]データクラス</paramref>
    ' ''' </summary>
    ' ''' <returns>boolean  取得合否　True:取得成功 Flase:取得失敗</returns>
    ' ''' <remarks>科目マスタより、データを取得
    ' ''' <para>作成情報：2015/08/13 yu.satoh
    ' ''' <p>改訂情報 : </p>
    ' ''' </para></remarks>
    'Public Function CmbKariShosaiSet(ByRef dataEXTM0102 As DataEXTM0102, ByRef karikamokuCd As String, ByRef kariSaimokuCd As String, ByRef kariUchiCd As String) As Boolean

    '    '開始ログ出力()
    '    CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

    '    '変数宣言
    '    Dim Cn As New NpgsqlConnection(DbString)    'コネクション
    '    Dim Adapter As New NpgsqlDataAdapter        'アダプタ
    '    Dim Table As New DataTable()                'テーブル

    '    Try

    '        'コネクションを開く
    '        Cn.Open()

    '        '付帯設備マスタメンテ用SQLの作成・設定
    '        If sqlEXTM0102.CmbKariShosai(Adapter, Cn, dataEXTM0102, karikamokuCd, kariSaimokuCd, kariUchiCd) = False Then
    '            Return False
    '        End If

    '        'SQLログ
    '        CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "システム管理情報", Nothing, Adapter.SelectCommand)

    '        'データを取得
    '        Adapter.Fill(Table)
    '        '取得データをデータクラスへ保存
    '        dataEXTM0102.PropDtKamokuMst = Table

    '        '終了ログ出力
    '        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

    '        Return True
    '    Catch ex As Exception
    '        '例外発生
    '        CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
    '        puErrMsg = EXTM0102_E0000 & ex.Message
    '        Return False
    '    Finally
    '        Cn.Close()
    '        Cn.Dispose()
    '        Adapter.Dispose()
    '        Table.Dispose()
    '    End Try

    'End Function
    ' 2016.04.28 DEL END↑ h.hagiwara レスポンス改善

    ''' <summary>
    ''' 付帯設備マスタメンテ、分類表のデータが増減した際、表示行数を再設定
    ''' <paramref name="dataEXTM0102">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <remarks>データのある行を確認し、その行+5行だけ表に表示
    ''' <para>作成情報：2015/08/17 yu.satoh
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Sub ChangeGroupingInsartRow(ByRef dataEXTM0102 As DataEXTM0102)

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Dim cnt As Integer = 0
        For i = 0 To dataEXTM0102.PropVwGroupingSheet.ActiveSheet.RowCount - 1
            For j = 0 To dataEXTM0102.PropVwGroupingSheet.ActiveSheet.ColumnCount - 1
                If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, j).Text <> Nothing Then
                    cnt = i + 1
                    Exit For
                End If
            Next
        Next

        'データのある行＋５行だけ行を表示
        dataEXTM0102.PropVwGroupingSheet.ActiveSheet.RowCount = cnt + 5

        'シート上コンボボックス設定
        For i = cnt To dataEXTM0102.PropVwGroupingSheet.ActiveSheet.RowCount - 1
            'i4とi8はここで全てコンボボックスとして宣言
            dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, M0102_BUNRUI_COL_KANJYO).CellType = cmbKamoku
            ' 2016.04.28 DEL START↓ h.hagiwara レスポンス改善
            'dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 8).CellType = cmbKarikamoku
            ' 2016.04.28 DEL END↑ h.hagiwara レスポンス改善
            '勘定科目コンボボックス作成
            CmbKamokuCdCreate(dataEXTM0102)
            ' 2016.04.28 DEL START↓ h.hagiwara レスポンス改善
            ''借方勘定科目コンボボックス作成
            'CmbKarikamokuCreate(dataEXTM0102)
            ' 2016.04.28 DEL END↑ h.hagiwara レスポンス改善
        Next

        '終了ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

    End Sub

    ''' <summary>
    ''' 付帯設備マスタメンテ、付帯設備表のデータが増減した際、表示行数を再設定
    ''' <paramref name="dataEXTM0102">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <remarks>データのある行を確認し、その行+5行だけ表に表示
    ''' <para>作成情報：2015/08/17 yu.satoh
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Sub ChangeFutaiInsartRow(ByRef dataEXTM0102 As DataEXTM0102)

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim cnt As Integer = 0
        Dim RowCnt As Integer

        'データのある行を検索
        For i = 0 To dataEXTM0102.PropVwFutaiSheet.ActiveSheet.RowCount - 1
            For j = 0 To dataEXTM0102.PropVwFutaiSheet.ActiveSheet.ColumnCount - 4
                If dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, j).Value <> Nothing Then
                    cnt = i + 1
                    RowCnt = dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, M0102_FUTAI_COL_ROW).Text
                    Exit For
                End If
            Next
        Next

        'データのある行＋５行を表示
        dataEXTM0102.PropVwFutaiSheet.ActiveSheet.RowCount = cnt + 5
        For j = cnt - 1 To cnt - 1 + 5
            dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(j, M0102_FUTAI_COL_ROW).Text = RowCnt
        Next

        '終了ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

    End Sub

    ''' <summary>
    ''' 付帯設備マスタメンテ、科目コードコンボボックス作成
    ''' <paramref name="dataEXTM0102">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <remarks>コンボボックス作成
    ''' <para>作成情報：2015/08/17 yu.satoh
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Sub CmbKamokuCdCreate(dataEXTM0102 As DataEXTM0102)

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        'リストに表示されるアイテムを定義します
        cmbKamoku.Items = New String() {""}
        cmbKamoku.ItemData = New String() {Nothing}

        ' 2015.11.25 ADD START↓ h.hagiwara
        Dim aryNm As ArrayList = New ArrayList()
        Dim aryData As ArrayList = New ArrayList()

        aryNm.AddRange(cmbKamoku.Items)
        aryData.AddRange(cmbKamoku.ItemData)
        ' 2015.11.25 ADD END↑ h.hagiwara

        For i = 0 To dataEXTM0102.PropDtKanzyoCd.Rows.Count - 1

            ' 2015.11.25 DEL START↓ h.hagiwara
            'Dim aryNm As ArrayList = New ArrayList()
            'Dim aryData As ArrayList = New ArrayList()

            'aryNm.AddRange(cmbKamoku.Items)
            'aryData.AddRange(cmbKamoku.ItemData)
            ' 2015.11.25 DEL END↑ h.hagiwara
            aryNm.Add(dataEXTM0102.PropDtKanzyoCd.Rows(i).Item(1).ToString)
            aryData.Add(dataEXTM0102.PropDtKanzyoCd.Rows(i).Item(0).ToString)

            ' 2015.11.25 DEL START↓ h.hagiwara
            'cmbKamoku.Items = CType(aryNm.ToArray(GetType(String)), String())
            'cmbKamoku.ItemData = CType(aryData.ToArray(GetType(String)), String())

            ''セルから取得／設定する値はItemDataとする
            'cmbKamoku.EditorValue = FarPoint.Win.Spread.CellType.EditorValue.ItemData
            ' 2015.11.25 DEL END↑ h.hagiwara
        Next

        ' 2015.11.25 ADD START↓ h.hagiwara
        cmbKamoku.Items = CType(aryNm.ToArray(GetType(String)), String())
        cmbKamoku.ItemData = CType(aryData.ToArray(GetType(String)), String())
        'セルから取得／設定する値はItemDataとする
        cmbKamoku.EditorValue = FarPoint.Win.Spread.CellType.EditorValue.ItemData
        ' 2015.11.25 ADD END↑ h.hagiwara

        '終了ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

    End Sub

    ''' <summary>
    ''' 付帯設備マスタメンテ、科目コードが変更された場合、細目コードを活性化し、コンボボックス作成
    ''' <paramref name="dataEXTM0102">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <remarks>細目コード活性化、コンボボックス作成
    ''' <para>作成情報：2015/08/17 yu.satoh
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Sub ChangeCmbKamokuCd(ByRef raw As Integer, dataEXTM0102 As DataEXTM0102)

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '細目コンボボックスを作成
        'If CmbSaimokuSet(dataEXTM0102, dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(raw, 4).Value) = True Then     ' 2016.06.24 UPD h.hagiwara コンボリスト設定方法変更対応
        If dataEXTM0102.PropDtSaimoku.Rows.Count > 0 Then                                                                 ' 2016.06.24 UPD h.hagiwara コンボリスト設定方法変更対応
            Dim cmbSaimoku As New FarPoint.Win.Spread.CellType.ComboBoxCellType()
            dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(raw, 5).CellType = cmbSaimoku

            cmbSaimoku.Items = New String() {""}
            cmbSaimoku.ItemData = New String() {Nothing}
            ' 2015.11.25 ADD START↓ h.hagiwara
            Dim aryNm As ArrayList = New ArrayList()
            Dim aryData As ArrayList = New ArrayList()

            aryNm.AddRange(cmbSaimoku.Items)
            aryData.AddRange(cmbSaimoku.ItemData)
            ' 2015.11.25 ADD END↑ h.hagiwara

            ' 2016.06.24 ADD START↓ h.hagiwara コンボリスト設定方法変更対応
            dataEXTM0102.PropStrSelKanjo = dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(raw, 4).Value
            Dim dr_Set() As DataRow = dataEXTM0102.PropDtSaimoku.Select(dataEXTM0102.PropDtSaimoku.Columns(0).Caption & "='" & dataEXTM0102.PropStrSelKanjo & "'")
            ' 2016.06.24 ADD END↑ h.hagiwara コンボリスト設定方法変更対応

            'For j = 0 To dataEXTM0102.PropDtKamokuMst.Rows.Count - 1                             ' 2016.06.24 UPD h.hagiwara コンボリスト設定方法変更対応
            For j = 0 To dr_Set.Length - 1                                                        ' 2016.06.24 UPD h.hagiwara コンボリスト設定方法変更対応
                ' 2015.11.25 DEL START↓ h.hagiwara
                'Dim aryNm As ArrayList = New ArrayList()
                'Dim aryData As ArrayList = New ArrayList()

                'aryNm.AddRange(cmbSaimoku.Items)
                'aryData.AddRange(cmbSaimoku.ItemData)
                ' 2015.11.25 DEL END↑ h.hagiwara
                ' 2016.06.24 UPD START↓ h.hagiwara コンボリスト設定方法変更対応
                'aryNm.Add(dataEXTM0102.PropDtKamokuMst.Rows(j).Item(1).ToString)
                'aryData.Add(dataEXTM0102.PropDtKamokuMst.Rows(j).Item(0).ToString)
                aryNm.Add(dr_Set(j).Item(2).ToString)
                aryData.Add(dr_Set(j).Item(1).ToString)
                ' 2016.06.24 UPD END↑ h.hagiwara コンボリスト設定方法変更対応
                ' 2015.11.25 DEL START↓ h.hagiwara
                'cmbSaimoku.Items = CType(aryNm.ToArray(GetType(String)), String())
                'cmbSaimoku.ItemData = CType(aryData.ToArray(GetType(String)), String())

                ''セルから取得／設定する値はItemDataとする
                'cmbSaimoku.EditorValue = FarPoint.Win.Spread.CellType.EditorValue.ItemData
                ' 2015.11.25 DEL END↑ h.hagiwara
            Next
            ' 2015.11.25 ADD START↓ h.hagiwara
            cmbSaimoku.Items = CType(aryNm.ToArray(GetType(String)), String())
            cmbSaimoku.ItemData = CType(aryData.ToArray(GetType(String)), String())
            'セルから取得／設定する値はItemDataとする
            cmbSaimoku.EditorValue = FarPoint.Win.Spread.CellType.EditorValue.ItemData
            ' 2015.11.25 ADD END↑ h.hagiwara
        End If

        '細目コードを初期化、活性化し、内訳、詳細コードを非活性にする
        dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(raw, 5).Locked = False
        dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(raw, 6).Value = ""
        dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(raw, 6).Locked = True
        dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(raw, 7).Value = ""
        dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(raw, 7).Locked = True

        '終了ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

    End Sub

    ''' <summary>
    ''' 付帯設備マスタメンテ、細目コードが変更された場合、内訳コードを活性化し、コンボボックス作成
    ''' <paramref name="dataEXTM0102">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <remarks>内訳コード活性化、コンボボックス作成
    ''' <para>作成情報：2015/08/17 yu.satoh
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Sub ChangeCmbSaimokuCd(ByRef raw As Integer, dataEXTM0102 As DataEXTM0102)

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        'If CmbUchiwakeSet(dataEXTM0102, dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(raw, 4).Value, _                                   ' 2016.06.24 UPD h.hagiwara コンボリスト設定方法変更対応
        '                  dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(raw, 5).Value) = True Then                                       ' 2016.06.24 UPD h.hagiwara コンボリスト設定方法変更対応
        If dataEXTM0102.PropDtUchiwake.Rows.Count > 0 Then                                                                                     ' 2016.06.24 UPD h.hagiwara コンボリスト設定方法変更対応
            'セルタイプと、コンボボックスの設定 
            Dim cmbUchi As New FarPoint.Win.Spread.CellType.ComboBoxCellType()
            dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(raw, 6).CellType = cmbUchi
            'セルから取得／設定する値はItemDataとする
            cmbUchi.EditorValue = FarPoint.Win.Spread.CellType.EditorValue.ItemData

            cmbUchi.Items = New String() {""}
            cmbUchi.ItemData = New String() {Nothing}
            ' 2015.11.25 ADD START↓ h.hagiwara
            Dim aryNm As ArrayList = New ArrayList()
            Dim aryData As ArrayList = New ArrayList()

            aryNm.AddRange(cmbUchi.Items)
            aryData.AddRange(cmbUchi.ItemData)
            ' 2015.11.25 ADD END↑ h.hagiwara

            ' 2016.06.24 ADD START↓ h.hagiwara コンボリスト設定方法変更対応
            dataEXTM0102.PropStrSelKanjo = dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(raw, 4).Value
            dataEXTM0102.PropStrSelSaimoku = dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(raw, 5).Value
            Dim dr_Set() As DataRow = dataEXTM0102.PropDtUchiwake.Select(dataEXTM0102.PropDtUchiwake.Columns(0).Caption & "='" & dataEXTM0102.PropStrSelKanjo & "' AND " & dataEXTM0102.PropDtUchiwake.Columns(1).Caption & "='" & dataEXTM0102.PropStrSelSaimoku & "'")
            ' 2016.06.24 ADD END↑ h.hagiwara コンボリスト設定方法変更対応

            'For j = 0 To dataEXTM0102.PropDtKamokuMst.Rows.Count - 1                             ' 2016.06.24 UPD h.hagiwara コンボリスト設定方法変更対応
            For j = 0 To dr_Set.Length - 1                                                        ' 2016.06.24 UPD h.hagiwara コンボリスト設定方法変更対応
                ' 2015.11.25 DEL START↓ h.hagiwara
                'Dim aryNm As ArrayList = New ArrayList()
                'Dim aryData As ArrayList = New ArrayList()

                'aryNm.AddRange(cmbUchi.Items)
                'aryData.AddRange(cmbUchi.ItemData)
                ' 2015.11.25 DEL END↑ h.hagiwara
                ' 2016.06.24 UPD START↓ h.hagiwara コンボリスト設定方法変更対応
                'aryNm.Add(dataEXTM0102.PropDtKamokuMst.Rows(j).Item(1).ToString)
                'aryData.Add(dataEXTM0102.PropDtKamokuMst.Rows(j).Item(0).ToString)
                aryNm.Add(dr_Set(j).Item(3).ToString)
                aryData.Add(dr_Set(j).Item(2).ToString)
                ' 2016.06.24 UPD END↑ h.hagiwara コンボリスト設定方法変更対応

                ' 2015.11.25 DEL START↓ h.hagiwara
                'cmbUchi.Items = CType(aryNm.ToArray(GetType(String)), String())
                'cmbUchi.ItemData = CType(aryData.ToArray(GetType(String)), String())
                ' 2015.11.25 DEL END↑ h.hagiwara
            Next
            ' 2015.11.25 ADD START↓ h.hagiwara
            cmbUchi.Items = CType(aryNm.ToArray(GetType(String)), String())
            cmbUchi.ItemData = CType(aryData.ToArray(GetType(String)), String())
            'セルから取得／設定する値はItemDataとする
            cmbUchi.EditorValue = FarPoint.Win.Spread.CellType.EditorValue.ItemData
            ' 2015.11.25 ADD END↑ h.hagiwara
        End If

        '内訳コードを活性化し、詳細コードを非活性にする
        dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(raw, 6).Value = ""
        dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(raw, 6).Locked = False
        dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(raw, 7).Value = ""
        dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(raw, 7).Locked = True

        '終了ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

    End Sub

    ''' <summary>
    ''' 付帯設備マスタメンテ、内訳コードが変更された場合、詳細コードを活性化し、コンボボックス作成
    ''' <paramref name="dataEXTM0102">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <remarks>詳細コード活性化、コンボボックス作成
    ''' <para>作成情報：2015/08/17 yu.satoh
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Sub ChangeCmbUchiwakeCd(ByRef raw As Integer, dataEXTM0102 As DataEXTM0102)

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(raw, 7).Locked = False

        'If CmbShosaiSet(dataEXTM0102, dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(raw, 4).Value, _                             ' 2016.06.24 UPD h.hagiwara コンボリスト設定方法変更対応
        '      dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(raw, 5).Value, _                                                     ' 2016.06.24 UPD h.hagiwara コンボリスト設定方法変更対応
        '      dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(raw, 6).Value) = True Then                                           ' 2016.06.24 UPD h.hagiwara コンボリスト設定方法変更対応
        If dataEXTM0102.PropDtShosai.Rows.Count > 0 Then                                                                               ' 2016.06.24 UPD h.hagiwara コンボリスト設定方法変更対応
            'セルタイプと、コンボボックスの設定
            Dim cmbShosai As New FarPoint.Win.Spread.CellType.ComboBoxCellType()
            dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(raw, 7).CellType = cmbShosai

            cmbShosai.Items = New String() {""}
            cmbShosai.ItemData = New String() {Nothing}
            ' 2015.11.25 ADD START↓ h.hagiwara
            Dim aryNm As ArrayList = New ArrayList()
            Dim aryData As ArrayList = New ArrayList()

            aryNm.AddRange(cmbShosai.Items)
            aryData.AddRange(cmbShosai.ItemData)
            ' 2015.11.25 ADD END↑ h.hagiwara

            ' 2016.06.24 ADD START↓ h.hagiwara コンボリスト設定方法変更対応
            dataEXTM0102.PropStrSelKanjo = dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(raw, 4).Value
            dataEXTM0102.PropStrSelSaimoku = dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(raw, 5).Value
            dataEXTM0102.PropStrSelUchiwake = dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(raw, 6).Value
            Dim dr_Set() As DataRow = dataEXTM0102.PropDtShosai.Select(dataEXTM0102.PropDtShosai.Columns(0).Caption & "='" & dataEXTM0102.PropStrSelKanjo & "' AND " & dataEXTM0102.PropDtShosai.Columns(1).Caption & "='" & dataEXTM0102.PropStrSelSaimoku & "' AND " & dataEXTM0102.PropDtShosai.Columns(2).Caption & "='" & dataEXTM0102.PropStrSelUchiwake & "'")
            ' 2016.06.24 ADD END↑ h.hagiwara コンボリスト設定方法変更対応

            'For j = 0 To dataEXTM0102.PropDtKamokuMst.Rows.Count - 1                             ' 2016.06.24 UPD h.hagiwara コンボリスト設定方法変更対応
            For j = 0 To dr_Set.Length - 1                                                        ' 2016.06.24 UPD h.hagiwara コンボリスト設定方法変更対応
                ' 2015.11.25 DEL START↓ h.hagiwara
                'Dim aryNm As ArrayList = New ArrayList()
                'Dim aryData As ArrayList = New ArrayList()

                'aryNm.AddRange(cmbShosai.Items)
                'aryData.AddRange(cmbShosai.ItemData)
                ' 2015.11.25 DEL END↑ h.hagiwara
                ' 2016.06.24 UPD START↓ h.hagiwara コンボリスト設定方法変更対応
                'aryNm.Add(dataEXTM0102.PropDtKamokuMst.Rows(j).Item(1).ToString)
                'aryData.Add(dataEXTM0102.PropDtKamokuMst.Rows(j).Item(0).ToString)
                aryNm.Add(dr_Set(j).Item(4).ToString)
                aryData.Add(dr_Set(j).Item(3).ToString)
                ' 2016.06.24 UPD END↑ h.hagiwara コンボリスト設定方法変更対応
                ' 2015.11.25 DEL START↓ h.hagiwara
                'cmbShosai.Items = CType(aryNm.ToArray(GetType(String)), String())
                'cmbShosai.ItemData = CType(aryData.ToArray(GetType(String)), String())

                ''セルから取得／設定する値はItemDataとする
                'cmbShosai.EditorValue = FarPoint.Win.Spread.CellType.EditorValue.ItemData
                ' 2015.11.25 DEL END↑ h.hagiwara
            Next
            ' 2015.11.25 ADD START↓ h.hagiwara
            cmbShosai.Items = CType(aryNm.ToArray(GetType(String)), String())
            cmbShosai.ItemData = CType(aryData.ToArray(GetType(String)), String())
            'セルから取得／設定する値はItemDataとする
            cmbShosai.EditorValue = FarPoint.Win.Spread.CellType.EditorValue.ItemData
            ' 2015.11.25 ADD END↑ h.hagiwara
        End If

        '終了ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

    End Sub

    ' 2016.04.28 DEL START↓ h.hagiwara レスポンス改善
    ' ''' <summary>
    ' ''' 付帯設備マスタメンテ、借方科目コードが変更された場合、借方細目コードを活性化し、コンボボックス作成
    ' ''' <paramref name="dataEXTM0102">[IN/OUT]データクラス</paramref>
    ' ''' </summary>
    ' ''' <remarks>借方細目コード活性化、コンボボックス作成
    ' ''' <para>作成情報：2015/08/17 yu.satoh
    ' ''' <p>改訂情報 : </p>
    ' ''' </para></remarks>
    'Public Sub ChangeCmbKarikamokuCd(ByRef raw As Integer, dataEXTM0102 As DataEXTM0102)

    '    '開始ログ出力()
    '    CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

    '    dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(raw, 9).Locked = False
    '    dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(raw, 9).Value = ""

    '    If CmbKariSaimokuSet(dataEXTM0102, dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(raw, 8).Value) = True Then
    '        'セルタイプと、コンボボックスの設定
    '        Dim cmbKarisaimoku As New FarPoint.Win.Spread.CellType.ComboBoxCellType()
    '        dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(raw, 9).CellType = cmbKarisaimoku

    '        cmbKarisaimoku.Items = New String() {""}
    '        cmbKarisaimoku.ItemData = New String() {Nothing}
    '        ' 2015.11.25 ADD START↓ h.hagiwara
    '        Dim aryNm As ArrayList = New ArrayList()
    '        Dim aryData As ArrayList = New ArrayList()

    '        aryNm.AddRange(cmbKarisaimoku.Items)
    '        aryData.AddRange(cmbKarisaimoku.ItemData)
    '        ' 2015.11.25 ADD END↑ h.hagiwara

    '        For j = 0 To dataEXTM0102.PropDtKamokuMst.Rows.Count - 1
    '            ' 2015.11.25 DEL START↓ h.hagiwara
    '            'Dim aryNm As ArrayList = New ArrayList()
    '            'Dim aryData As ArrayList = New ArrayList()

    '            'aryNm.AddRange(cmbKarisaimoku.Items)
    '            'aryData.AddRange(cmbKarisaimoku.ItemData)
    '            ' 2015.11.25 DEL END↑ h.hagiwara
    '            aryNm.Add(dataEXTM0102.PropDtKamokuMst.Rows(j).Item(1).ToString)
    '            aryData.Add(dataEXTM0102.PropDtKamokuMst.Rows(j).Item(0).ToString)
    '            ' 2015.11.25 DEL START↓ h.hagiwara
    '            'cmbKarisaimoku.Items = CType(aryNm.ToArray(GetType(String)), String())
    '            'cmbKarisaimoku.ItemData = CType(aryData.ToArray(GetType(String)), String())

    '            ''セルから取得／設定する値はItemDataとする
    '            'cmbKarisaimoku.EditorValue = FarPoint.Win.Spread.CellType.EditorValue.ItemData
    '            ' 2015.11.25 DEL END↑ h.hagiwara
    '        Next
    '        ' 2015.11.25 ADD START↓ h.hagiwara
    '        cmbKarisaimoku.Items = CType(aryNm.ToArray(GetType(String)), String())
    '        cmbKarisaimoku.ItemData = CType(aryData.ToArray(GetType(String)), String())
    '        'セルから取得／設定する値はItemDataとする
    '        cmbKarisaimoku.EditorValue = FarPoint.Win.Spread.CellType.EditorValue.ItemData
    '        ' 2015.11.25 ADD END↑ h.hagiwara
    '    End If

    '    '借方内訳、借方詳細を非活性にする
    '    dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(raw, 10).Locked = True
    '    dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(raw, 10).Value = ""
    '    dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(raw, 11).Locked = True
    '    dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(raw, 11).Value = ""

    '    '終了ログ出力
    '    CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

    'End Sub

    ' ''' <summary>
    ' ''' 付帯設備マスタメンテ、借方細目コードが変更された場合、借方内訳コードを活性化し、コンボボックス作成
    ' ''' <paramref name="dataEXTM0102">[IN/OUT]データクラス</paramref>
    ' ''' </summary>
    ' ''' <remarks>借方内訳コード活性化、コンボボックス作成
    ' ''' <para>作成情報：2015/08/17 yu.satoh
    ' ''' <p>改訂情報 : </p>
    ' ''' </para></remarks>
    'Public Sub ChangeCmbKarisaimokuCd(ByRef raw As Integer, dataEXTM0102 As DataEXTM0102)

    '    '開始ログ出力()
    '    CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

    '    dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(raw, 10).Locked = False
    '    dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(raw, 10).Value = ""

    '    '借方内訳
    '    If CmbKariUchiSet(dataEXTM0102, dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(raw, 8).Value, _
    '                     dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(raw, 9).Value) = True Then
    '        'セルタイプと、コンボボックスの設定
    '        Dim cmbKariuchi As New FarPoint.Win.Spread.CellType.ComboBoxCellType()
    '        dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(raw, 10).CellType = cmbKariuchi
    '        cmbKariuchi.Items = New String() {""}
    '        cmbKariuchi.ItemData = New String() {Nothing}
    '        ' 2015.11.25 ADD START↓ h.hagiwara
    '        Dim aryNm As ArrayList = New ArrayList()
    '        Dim aryData As ArrayList = New ArrayList()

    '        aryNm.AddRange(cmbKariuchi.Items)
    '        aryData.AddRange(cmbKariuchi.ItemData)
    '        ' 2015.11.25 ADD END↑ h.hagiwara

    '        For j = 0 To dataEXTM0102.PropDtKamokuMst.Rows.Count - 1
    '            ' 2015.11.25 DEL START↓ h.hagiwara
    '            'Dim aryNm As ArrayList = New ArrayList()
    '            'Dim aryData As ArrayList = New ArrayList()

    '            'aryNm.AddRange(cmbKariuchi.Items)
    '            'aryData.AddRange(cmbKariuchi.ItemData)
    '            ' 2015.11.25 DEL END↑ h.hagiwara
    '            aryNm.Add(dataEXTM0102.PropDtKamokuMst.Rows(j).Item(1).ToString)
    '            aryData.Add(dataEXTM0102.PropDtKamokuMst.Rows(j).Item(0).ToString)
    '            ' 2015.11.25 DEL START↓ h.hagiwara
    '            'cmbKariuchi.Items = CType(aryNm.ToArray(GetType(String)), String())
    '            'cmbKariuchi.ItemData = CType(aryData.ToArray(GetType(String)), String())

    '            ''セルから取得／設定する値はItemDataとする
    '            'cmbKariuchi.EditorValue = FarPoint.Win.Spread.CellType.EditorValue.ItemData
    '            ' 2015.11.25 DEL END↑ h.hagiwara
    '        Next
    '        ' 2015.11.25 ADD START↓ h.hagiwara
    '        cmbKariuchi.Items = CType(aryNm.ToArray(GetType(String)), String())
    '        cmbKariuchi.ItemData = CType(aryData.ToArray(GetType(String)), String())
    '        'セルから取得／設定する値はItemDataとする
    '        cmbKariuchi.EditorValue = FarPoint.Win.Spread.CellType.EditorValue.ItemData
    '        ' 2015.11.25 ADD END↑ h.hagiwara
    '    End If

    '    '借方詳細を非活性に
    '    dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(raw, 11).Value = ""
    '    dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(raw, 11).Locked = True

    '    '終了ログ出力
    '    CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

    'End Sub

    ' ''' <summary>
    ' ''' 付帯設備マスタメンテ、借方内訳コードが変更された場合、借方詳細コードを活性化し、コンボボックス作成
    ' ''' <paramref name="dataEXTM0102">[IN/OUT]データクラス</paramref>
    ' ''' </summary>
    ' ''' <remarks>借方詳細コード活性化、コンボボックス作成
    ' ''' <para>作成情報：2015/08/17 yu.satoh
    ' ''' <p>改訂情報 : </p>
    ' ''' </para></remarks>
    'Public Sub ChangeCmbKariuchiwakeCd(ByRef raw As Integer, dataEXTM0102 As DataEXTM0102)

    '    '開始ログ出力()
    '    CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

    '    dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(raw, 11).Value = ""
    '    dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(raw, 11).Locked = False

    '    '借方詳細
    '    If CmbKariShosaiSet(dataEXTM0102, dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(raw, 8).Value, _
    '                      dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(raw, 9).Value, _
    '                      dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(raw, 10).Value) = True Then

    '        'セルタイプと、コンボボックスの設定
    '        Dim cmbKarisyosai As New FarPoint.Win.Spread.CellType.ComboBoxCellType()
    '        dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(raw, 11).CellType = cmbKarisyosai

    '        cmbKarisyosai.Items = New String() {""}
    '        cmbKarisyosai.ItemData = New String() {Nothing}
    '        ' 2015.11.25 ADD START↓ h.hagiwara
    '        Dim aryNm As ArrayList = New ArrayList()
    '        Dim aryData As ArrayList = New ArrayList()

    '        aryNm.AddRange(cmbKarisyosai.Items)
    '        aryData.AddRange(cmbKarisyosai.ItemData)
    '        ' 2015.11.25 ADD END↑ h.hagiwara

    '        For j = 0 To dataEXTM0102.PropDtKamokuMst.Rows.Count - 1
    '            ' 2015.11.25 DEL START↓ h.hagiwara
    '            'Dim aryNm As ArrayList = New ArrayList()
    '            'Dim aryData As ArrayList = New ArrayList()

    '            'aryNm.AddRange(cmbKarisyosai.Items)
    '            'aryData.AddRange(cmbKarisyosai.ItemData)
    '            ' 2015.11.25 DEL END↑ h.hagiwara
    '            aryNm.Add(dataEXTM0102.PropDtKamokuMst.Rows(j).Item(1).ToString)
    '            aryData.Add(dataEXTM0102.PropDtKamokuMst.Rows(j).Item(0).ToString)
    '            ' 2015.11.25 DEL START↓ h.hagiwara
    '            'cmbKarisyosai.Items = CType(aryNm.ToArray(GetType(String)), String())
    '            'cmbKarisyosai.ItemData = CType(aryData.ToArray(GetType(String)), String())

    '            ''セルから取得／設定する値はItemDataとする
    '            'cmbKarisyosai.EditorValue = FarPoint.Win.Spread.CellType.EditorValue.ItemData
    '            ' 2015.11.25 DEL END↑ h.hagiwara
    '        Next
    '        ' 2015.11.25 ADD START↓ h.hagiwara
    '        cmbKarisyosai.Items = CType(aryNm.ToArray(GetType(String)), String())
    '        cmbKarisyosai.ItemData = CType(aryData.ToArray(GetType(String)), String())
    '        'セルから取得／設定する値はItemDataとする
    '        cmbKarisyosai.EditorValue = FarPoint.Win.Spread.CellType.EditorValue.ItemData
    '        ' 2015.11.25 ADD END↑ h.hagiwara
    '    End If

    '    '終了ログ出力
    '    CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

    'End Sub
    ' 2016.04.28 DEL END↑ h.hagiwara レスポンス改善

    ''' <summary>
    ''' 付帯設備マスタメンテ、登録ボタン処理
    ''' <paramref name="dataEXTM0102">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <remarks>登録ボタン押下時、分類表の項目をインサート、もしくはアップデート処理
    ''' <para>作成情報：2015/08/18 yu.satoh
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function EntryBunrui(ByRef i As Integer, dataEXTM0102 As DataEXTM0102, _
                                ByRef noData As Integer, ByRef Tsx As NpgsqlTransaction, _
                                          ByVal Cn As NpgsqlConnection) As Boolean

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        'Dim Cn As New NpgsqlConnection(DbString)    'コネクション
        Dim Adapter As New NpgsqlCommand        'アダプタ
        Dim Table As New DataTable()                'テーブル

        Try

            'コネクションを開く
            'Cn.Open()

            If dataEXTM0102.PropNewBtn.Checked = True Then
                ' 最新値を取得する
                If GetFutaiBunruicd(dataEXTM0102) = False Then
                    Return False
                End If
                '値がなければ挿入.ただし、その行全てが空の場合、挿入しない
                If sqlEXTM0102.InsertBunrui(i, Adapter, Tsx, Cn, dataEXTM0102) = False Then
                    Return False
                End If
            Else
                '分類表、値があれば更新
                If dataEXTM0102.PropDtFbunruiMst.Rows.Count >= i + 1 Then
                    If sqlEXTM0102.UpdateBunrui(i, Adapter, Tsx, Cn, dataEXTM0102) = False Then
                        Return False
                    End If
                Else
                    '' 最新値を取得する
                    'If GetBunruicd(dataEXTM0102) = False Then
                    '    Return False
                    'End If
                    ' 最新値を取得する
                    If GetFutaiBunruicd(dataEXTM0102) = False Then
                        Return False
                    End If
                    '値がなければ挿入.ただし、その行全てが空の場合、挿入しない
                    If sqlEXTM0102.InsertBunrui(i, Adapter, Tsx, Cn, dataEXTM0102) = False Then
                        Return False
                    End If
                End If
            End If

            'SQLログ
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "システム管理情報", Nothing, Adapter)

            'コマンド実行
            Adapter.ExecuteNonQuery()

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            Return True

        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            'ロールバック
            If Tsx IsNot Nothing Then
                Tsx.Rollback()
            End If
            puErrMsg = EXTM0102_E0000 & ex.Message
            Return False
        Finally
            'Cn.Close()
            'Cn.Dispose()
            Adapter.Dispose()
            Table.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' 付帯設備マスタメンテ、登録ボタン処理
    ''' <paramref name="dataEXTM0102">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <remarks>登録ボタン押下時、付帯設備表の項目をインサート、もしくはアップデート処理
    ''' <para>作成情報：2015/08/18 yu.satoh
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function EntryFutai(dataEXTM0102 As DataEXTM0102 _
                               , ByRef Tsx As NpgsqlTransaction, _
                                          ByVal Cn As NpgsqlConnection) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        'Dim Cn As New NpgsqlConnection(DbString)    'コネクション
        Dim Adapter As New NpgsqlCommand        'アダプタ
        Dim Table As New DataTable()                'テーブル

        Try

            'コネクションを開く
            'Cn.Open()
            'If dataEXTM0102.PropNewBtn.Checked = True Then
            '    '値がなければ挿入,ただし、その行全てが空の場合、挿入しない
            '    If sqlEXTM0102.InsertFutai(i, Adapter, Tsx, Cn, dataEXTM0102) = False Then
            '        Return False
            '    End If
            'Else
            '    '付帯設備表、値があれば更新
            '    If dataEXTM0102.PropDtFutaiMst.Rows.Count >= i + 1 Then
            '        If sqlEXTM0102.UpdateFutai(i, Adapter, Tsx, Cn, dataEXTM0102) = False Then
            '            Return False
            '        End If
            '    Else
            '        '値がなければ挿入,ただし、その行全てが空の場合、挿入しない
            '        If sqlEXTM0102.InsertFutai(i, Adapter, Tsx, Cn, dataEXTM0102) = False Then
            '            Return False
            '        End If
            '    End If
            'End If

            'SQLログ
            dataEXTM0102.PropRow = -1
            If dataEXTM0102.PropNewBtn.Checked = True Then
                For j = 0 To dataEXTM0102.PropDtFutaiMst.Rows.Count - 1
                    If dataEXTM0102.PropRow = dataEXTM0102.PropDtFutaiMst.Rows(j).Item(M0102_FUTAI_COL_ROW) Then
                    Else
                        dataEXTM0102.PropRow = dataEXTM0102.PropDtFutaiMst.Rows(j).Item(M0102_FUTAI_COL_ROW)
                        dataEXTM0102.PropBunruiCd = dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(dataEXTM0102.PropRow, M0102_BUNRUI_COL_BUNRUICD).Value
                        dataEXTM0102.PropMaxFutaiBunruiCd = 0
                    End If

                    If sqlEXTM0102.InsertFutai(j, Adapter, Tsx, Cn, dataEXTM0102) = False Then
                        Return False
                    End If
                Next
            Else
                For j = 0 To dataEXTM0102.PropDtFutaiMst.Rows.Count - 1
                    If dataEXTM0102.PropRow = dataEXTM0102.PropDtFutaiMst.Rows(j).Item(M0102_FUTAI_COL_ROW) Then
                    Else
                        dataEXTM0102.PropRow = dataEXTM0102.PropDtFutaiMst.Rows(j).Item(M0102_FUTAI_COL_ROW)
                        dataEXTM0102.PropBunruiCd = dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(dataEXTM0102.PropRow, M0102_BUNRUI_COL_BUNRUICD).Value
                        dataEXTM0102.PropMaxFutaiBunruiCd = 0
                    End If

                    If dataEXTM0102.PropDtFutaiMst.Rows(j).Item(M0102_FUTAI_COL_UPKBN) = "0" Then
                        If sqlEXTM0102.UpdateFutai(j, Adapter, Tsx, Cn, dataEXTM0102) = False Then
                            Return False
                        End If
                        '2016.04.14 UPDATE START↓ h.hagiwara
                        'dataEXTM0102.PropMaxFutaiBunruiCd = dataEXTM0102.PropDtFutaiMst.Rows(j).Item(M0102_FUTAI_COL_SETUBICD)
                        If dataEXTM0102.PropMaxFutaiBunruiCd < dataEXTM0102.PropDtFutaiMst.Rows(j).Item(M0102_FUTAI_COL_SETUBICD) Then
                            dataEXTM0102.PropMaxFutaiBunruiCd = dataEXTM0102.PropDtFutaiMst.Rows(j).Item(M0102_FUTAI_COL_SETUBICD)
                        End If
                        '2016.04.14 UPDATE END↑ h.hagiwara
                    ElseIf dataEXTM0102.PropDtFutaiMst.Rows(j).Item(M0102_FUTAI_COL_UPKBN) = "1" Then
                        If sqlEXTM0102.InsertFutai(j, Adapter, Tsx, Cn, dataEXTM0102) = False Then
                            Return False
                        End If
                    End If

                Next
            End If
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "システム管理情報", Nothing, Adapter)

            ''コマンド実行
            'Adapter.ExecuteNonQuery()

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            Return True

        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            'ロールバック
            If Tsx IsNot Nothing Then
                Tsx.Rollback()
            End If
            puErrMsg = EXTM0102_E0000 & ex.Message
            Return False
        Finally
            'Cn.Close()
            'Cn.Dispose()
            Adapter.Dispose()
            Table.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' 付帯設備マスタメンテ、期間入力チェック
    ''' <paramref name="dataEXTM0102">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <remarks>登録ボタン押下時必須入力項目未記入、及びフォーマットに合わない場合にエラーメッセージ表示
    ''' <para>作成情報：2015/08/18 yu.satoh
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function CheckErrerKikan(dataEXTM0102 As DataEXTM0102) As Boolean

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim cnt As Integer = 0  'エラーチェック用変数

        '画面入力 from to
        Dim strKikanFrom As String = dataEXTM0102.PropYearFrom.Text + dataEXTM0102.PropMonthFrom.Text
        Dim strkikanTo As String = dataEXTM0102.PropYearTo.Text + dataEXTM0102.PropMonthTo.Text

        'コンボボックス内日付from to
        Dim strFkikanFrom As String '= dataEXTM0102.PropDtFinishedFromTo.Rows(0).Item(1).ToString.Substring(0, 4) + _
        'dataEXTM0102.PropDtFinishedFromTo.Rows(0).Item(1).ToString.Substring(5, 2)
        Dim strFkikanTo As String ' = dataEXTM0102.PropDtFinishedFromTo.Rows(0).Item(1).ToString.Substring(11, 4) + _
        'dataEXTM0102.PropDtFinishedFromTo.Rows(0).Item(1).ToString.Substring(16, 2)



        '期間from（年）入力チェック
        If dataEXTM0102.PropYearFrom.Text = Nothing Then
            puErrMsg = String.Format(EXTM0102_E0001, "期間(FROM)・年")
            Return False
        End If

        '期間from(年)が数字かどうか
        If Not IsNumeric(dataEXTM0102.PropYearFrom.Text) Then
            puErrMsg = String.Format(EXTM0102_E0003, "期間(FROM)・年")
            Return False
        End If

        '期間from（月）入力チェック
        If dataEXTM0102.PropMonthFrom.Text = Nothing Then
            puErrMsg = String.Format(EXTM0102_E0001, "期間(FROM)・月")
            Return False
        End If
        '期間from(月)が数字かどうか
        If Not IsNumeric(dataEXTM0102.PropMonthFrom.Text) Then
            puErrMsg = String.Format(EXTM0102_E0003, "期間(FROM)・月")
            Return False
        End If
        '期間from(月)入力チェック
        If CInt(dataEXTM0102.PropMonthFrom.Text) < 1 Or CInt(dataEXTM0102.PropMonthFrom.Text) > 12 Then
            puErrMsg = String.Format(EXTM0102_E0018, "期間(FROM)・月", "1", "12")
            Return False
        End If

        '期間to（年）入力チェック
        If dataEXTM0102.PropYearTo.Text = Nothing Then
            puErrMsg = String.Format(EXTM0102_E0001, "期間(TO)・年")
            Return False
        End If
        '期間To(年)が数字かどうか
        If Not IsNumeric(dataEXTM0102.PropYearTo.Text) Then
            puErrMsg = String.Format(EXTM0102_E0003, "期間(TO)・年")
            Return False
        End If

        '期間to（月）入力チェック
        If dataEXTM0102.PropMonthTo.Text = Nothing Then
            puErrMsg = String.Format(EXTM0102_E0001, "期間(TO)・月")
            Return False
        End If
        '期間To(月)が数字かどうか
        If Not IsNumeric(dataEXTM0102.PropMonthTo.Text) Then
            puErrMsg = String.Format(EXTM0102_E0003, "期間(TO)・月")
            Return False
        End If
        '期間from(月)入力チェック
        If CInt(dataEXTM0102.PropMonthTo.Text) < 1 Or CInt(dataEXTM0102.PropMonthTo.Text) > 12 Then
            puErrMsg = String.Format(EXTM0102_E0018, "期間(TO)・月", "1", "12")
            Return False
        End If

        ''期間from(年)が数字かどうか
        'If Not IsNumeric(dataEXTM0102.PropYearFrom.Text) Then
        '    puErrMsg = String.Format(EXTM0102_E0003, "期間(FROM)・年")
        '    Return False
        'End If

        ''期間from(月)が数字かどうか
        'If Not IsNumeric(dataEXTM0102.PropMonthFrom.Text) Then
        '    puErrMsg = String.Format(EXTM0102_E0003, "期間(FROM)・月")
        '    Return False
        'End If

        ''期間To(年)が数字かどうか
        'If Not IsNumeric(dataEXTM0102.PropYearTo.Text) Then
        '    puErrMsg = String.Format(EXTM0102_E0003, "期間(TO)・年")
        '    Return False
        'End If

        ''期間To(月)が数字かどうか
        'If Not IsNumeric(dataEXTM0102.PropMonthTo.Text) Then
        '    puErrMsg = String.Format(EXTM0102_E0003, "期間(TO)・月")
        '    Return False
        'End If

        ''期間from(月)入力チェック
        'If CInt(dataEXTM0102.PropMonthFrom.Text) < 1 Or CInt(dataEXTM0102.PropMonthFrom.Text) > 12 Then
        '    puErrMsg = String.Format(EXTM0102_E0018, "期間(FROM)・月", "1", "12")
        '    Return False
        'End If

        ''期間from(月)入力チェック
        'If CInt(dataEXTM0102.PropMonthTo.Text) < 1 Or CInt(dataEXTM0102.PropMonthTo.Text) > 12 Then
        '    puErrMsg = String.Format(EXTM0102_E0018, "期間(TO)・月", "1", "12")
        '    Return False
        'End If

        '期間逆転入力チェック
        If CInt(strkikanTo) - CInt(strKikanFrom) < 0 Then
            puErrMsg = String.Format(EXTM0102_E2011, "期間(TO)", "期間(FROM)")
            Return False
        End If

        '登録済み期間の分だけチェックを行う
        For i = 0 To dataEXTM0102.PropDtFinishedFromTo.Rows.Count - 1
            If dataEXTM0102.PropDtFinishedFromTo.Rows(i).Item(1) = Nothing Then
            Else
                'コンボボックス内の期間fromToを取得
                strFkikanFrom = dataEXTM0102.PropDtFinishedFromTo.Rows(i).Item(1).ToString.Substring(0, 4) + _
                                 dataEXTM0102.PropDtFinishedFromTo.Rows(i).Item(1).ToString.Substring(5, 2)
                strFkikanTo = dataEXTM0102.PropDtFinishedFromTo.Rows(i).Item(1).ToString.Substring(11, 4) + _
                         dataEXTM0102.PropDtFinishedFromTo.Rows(i).Item(1).ToString.Substring(16, 2)
                ' 現在表示期間（選択期間）は除く
                If strFkikanFrom = dataEXTM0102.PropFromYearUp + dataEXTM0102.PropFromMonthUp And _
                    strFkikanTo = dataEXTM0102.PropToYearUp + dataEXTM0102.PropToMonthUp Then
                Else
                    'エラーチェック用変数を初期化
                    cnt = 0
                    '入力期間が、設定済み期間よりも両方前ならカウントを加算
                    If CInt(strKikanFrom) < CInt(strFkikanFrom) And CInt(strkikanTo) < CInt(strFkikanFrom) Then
                        cnt += 1
                    End If
                    '入力期間が、設定済み期間よりも両方後ならカウントを加算
                    If CInt(strKikanFrom) > CInt(strFkikanTo) And CInt(strkikanTo) > CInt(strFkikanTo) Then
                        cnt += 1
                    End If
                    'もしカウントが一度もされなければエラーを表示
                    If cnt = 0 Then
                        puErrMsg = String.Format(EXTM0102_E2032)
                        Return False
                    End If
                End If
                ''エラーチェック用変数を初期化
                'cnt = 0
                ''入力期間が、設定済み期間よりも両方前ならカウントを加算
                'If CInt(strKikanFrom) < CInt(strFkikanFrom) And CInt(strkikanTo) < CInt(strFkikanFrom) Then
                '    cnt += 1
                'End If
                ''入力期間が、設定済み期間よりも両方後ならカウントを加算
                'If CInt(strKikanFrom) > CInt(strFkikanTo) And CInt(strkikanTo) > CInt(strFkikanTo) Then
                '    cnt += 1
                'End If
                ''もしカウントが一度もされなければエラーを表示
                'If cnt = 0 Then
                '    puErrMsg = String.Format(EXTM0102_E2032)
                '    Return False
                'End If
            End If
        Next

        ''期間逆転入力チェック
        'If CInt(strkikanTo) - CInt(strKikanFrom) < 0 Then
        '    puErrMsg = String.Format(EXTM0102_E2011, "期間(TO)", "期間(FROM)")
        '    Return False
        'End If

        '終了ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

        Return True

    End Function

    ''' <summary>
    ''' 付帯設備マスタメンテ、分類表入力チェック
    ''' <paramref name="dataEXTM0102">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <remarks>登録ボタン押下時必須入力項目未記入、及びフォーマットに合わない場合にエラーメッセージ表示
    ''' <para>作成情報：2015/08/18 yu.satoh
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function CheckErrerBunrui(dataEXTM0102 As DataEXTM0102) As Boolean

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '分類シート入力チェック
        For i = 0 To dataEXTM0102.PropVwGroupingSheet.ActiveSheet.RowCount - 1
            Dim errCnt As Integer = 0
            Dim errCnt2 As Integer = 0
            ''分類コード
            'For j = 1 To dataEXTM0102.PropVwGroupingSheet.ActiveSheet.ColumnCount - 1
            '    If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 0).Value = Nothing And dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, j).Value <> Nothing Then
            '        errCnt += 1
            '        Exit For
            '    End If
            'Next
            ''エラーがあればfalse返す
            'If errCnt <> 0 Then
            '    puErrMsg = String.Format(EXTM0102_E0001, "分類コード")
            '    Return False
            'End If
            ''分類コード2文字以内
            'If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 0).Text.Length > 2 Then
            '    puErrMsg = String.Format(EXTM0102_E0010, "分類コード", "2")
            '    Return False
            'End If

            '分類名
            'For j = 0 To dataEXTM0102.PropVwGroupingSheet.ActiveSheet.ColumnCount - 1
            '    'If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 1).Value = Nothing And dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, j).Value <> Nothing Then
            '    '    errCnt += 1
            '    '    Exit For
            '    'End If
            'Next
            If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 1).Value = Nothing Then
                For j = 1 To dataEXTM0102.PropVwGroupingSheet.ActiveSheet.ColumnCount - 1
                    If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, j).Value <> Nothing Then
                        errCnt += 1
                        Exit For
                    End If
                    If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, j).Value = Nothing And i = 0 Then
                        errCnt2 += 1
                        Exit For
                    End If
                Next
            End If
            'エラーがあればfalse返す
            If errCnt <> 0 Or errCnt2 <> 0 Then
                puErrMsg = String.Format(EXTM0102_E0001, "分類名")
                Return False
            End If
            '分類名20文字以内
            If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, M0102_BUNRUI_COL_BUNRUINM).Text.Length > 20 Then
                puErrMsg = String.Format(EXTM0102_E0010, "分類名", "20")
                Return False
            End If

            '集計キー
            'For j = 0 To dataEXTM0102.PropVwGroupingSheet.ActiveSheet.ColumnCount - 1
            '    If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 2).Value = Nothing And dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, j).Value <> Nothing Then
            '        errCnt += 1
            '        Exit For
            '    End If
            'Next
            If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, M0102_BUNRUI_COL_SYUKEIKY).Value = Nothing Then
                For j = 1 To dataEXTM0102.PropVwGroupingSheet.ActiveSheet.ColumnCount - 1
                    If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, j).Value <> Nothing Then
                        errCnt += 1
                        Exit For
                    End If
                Next
            End If
            'エラーがあればfalse返す
            If errCnt <> 0 Then
                puErrMsg = String.Format(EXTM0102_E0001, "集計キー")
                Return False
            End If
            '集計は20文字以内
            If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, M0102_BUNRUI_COL_SYUKEIKY).Text.Length > 20 Then
                puErrMsg = String.Format(EXTM0102_E0010, "集計キー", "20")
                Return False
            End If

            '勘定科目コード
            'For j = 0 To dataEXTM0102.PropVwGroupingSheet.ActiveSheet.ColumnCount - 1
            '    If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 4).Value = Nothing And dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, j).Value <> Nothing Then
            '        errCnt += 1
            '        Exit For
            '    End If
            'Next
            If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, M0102_BUNRUI_COL_KANJYO).Value = Nothing Then
                For j = 1 To dataEXTM0102.PropVwGroupingSheet.ActiveSheet.ColumnCount - 1
                    If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, j).Value <> Nothing Then
                        errCnt += 1
                        Exit For
                    End If
                Next
            End If
            'エラーがあればfalse返す
            If errCnt <> 0 Then
                puErrMsg = String.Format(EXTM0102_E0002, "勘定科目コード")
                Return False
            End If

            '細目コード
            'For j = 0 To dataEXTM0102.PropVwGroupingSheet.ActiveSheet.ColumnCount - 1
            '    If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 5).Value = Nothing And dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, j).Value <> Nothing Then
            '        errCnt += 1
            '        Exit For
            '    End If
            'Next
            If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, M0102_BUNRUI_COL_SAIMOKU).Value = Nothing Then
                For j = 1 To dataEXTM0102.PropVwGroupingSheet.ActiveSheet.ColumnCount - 1
                    If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, j).Value <> Nothing Then
                        errCnt += 1
                        Exit For
                    End If
                Next
            End If
            'エラーがあればfalse返す
            If errCnt <> 0 Then
                puErrMsg = String.Format(EXTM0102_E0002, "細目コード")
                Return False
            End If

            '内訳コード
            'For j = 0 To dataEXTM0102.PropVwGroupingSheet.ActiveSheet.ColumnCount - 1
            '    If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 6).Value = Nothing And dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, j).Value <> Nothing Then
            '        errCnt += 1
            '        Exit For
            '    End If
            'Next
            If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, M0102_BUNRUI_COL_UCHIWAKE).Value = Nothing Then
                For j = 1 To dataEXTM0102.PropVwGroupingSheet.ActiveSheet.ColumnCount - 1
                    If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, j).Value <> Nothing Then
                        errCnt += 1
                        Exit For
                    End If
                Next
            End If
            'エラーがあればfalse返す
            If errCnt <> 0 Then
                puErrMsg = String.Format(EXTM0102_E0002, "内訳コード")
                Return False
            End If

            '細目コード
            'For j = 0 To dataEXTM0102.PropVwGroupingSheet.ActiveSheet.ColumnCount - 1
            '    If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 7).Value = Nothing And dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, j).Value <> Nothing Then
            '        errCnt += 1
            '        Exit For
            '    End If
            'Next
            If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, M0102_BUNRUI_COL_SYOSAI).Value = Nothing Then
                For j = 1 To dataEXTM0102.PropVwGroupingSheet.ActiveSheet.ColumnCount - 1
                    If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, j).Value <> Nothing Then
                        errCnt += 1
                        Exit For
                    End If
                Next
            End If
            'エラーがあればfalse返す
            If errCnt <> 0 Then
                puErrMsg = String.Format(EXTM0102_E0002, "詳細コード")
                Return False
            End If

            '並び順
            'For j = 0 To dataEXTM0102.PropVwGroupingSheet.ActiveSheet.ColumnCount - 1
            '    If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 12).Value = Nothing And dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, j).Value <> Nothing Then
            '        errCnt += 1
            '        Exit For
            '    End If
            'Next
            If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, M0102_BUNRUI_COL_SORT).Value = Nothing Then
                For j = 0 To dataEXTM0102.PropVwGroupingSheet.ActiveSheet.ColumnCount - 1
                    If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, j).Value <> Nothing Then
                        errCnt += 1
                        Exit For
                    End If
                Next
            End If
            'エラーがあればfalse返す
            If errCnt <> 0 Then
                puErrMsg = String.Format(EXTM0102_E0001, "並び順")
                Return False
            End If

            If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, M0102_BUNRUI_COL_SORT).Value <> Nothing Then
                If ChkZenkaku(dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, M0102_BUNRUI_COL_SORT).Value) = False Then
                    puErrMsg = String.Format(EXTM0102_E0003, "並び順")
                    Return False
                End If
                '並び順は半角数字のみ
                If Not IsNumeric(dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, M0102_BUNRUI_COL_SORT).Value) Then
                    puErrMsg = String.Format(EXTM0102_E0003, "並び順")
                    Return False
                End If

                '並び順は4文字以内
                If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, M0102_BUNRUI_COL_SORT).Text.Length > 4 Then
                    puErrMsg = String.Format(EXTM0102_E0010, "並び順", "4")
                    Return False
                End If
            End If

        Next

        '終了ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

        Return True

    End Function

    ''' <summary>
    ''' 付帯設備マスタメンテ、付帯設備表入力チェック
    ''' <paramref name="dataEXTM0102">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <remarks>登録ボタン押下時必須入力項目未記入、及びフォーマットに合わない場合にエラーメッセージ表示
    ''' <para>作成情報：2015/08/18 yu.satoh
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function CheckErrerFutai(dataEXTM0102 As DataEXTM0102) As Boolean

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '付帯設備シート入力チェック
        For i = 0 To dataEXTM0102.PropVwFutaiSheet.ActiveSheet.RowCount - 1
            Dim errCnt As Integer = 0
            Dim errCnt2 As Integer = 0
            ''付帯設備コード
            'For j = 1 To dataEXTM0102.PropVwFutaiSheet.ActiveSheet.ColumnCount - 1
            '    If dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, 0).Value = Nothing And dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, j).Value <> Nothing Then
            '        errCnt += 1
            '        Exit For
            '    End If
            'Next
            ''エラーがあればfalse返す
            'If errCnt <> 0 Then
            '    puErrMsg = String.Format(EXTM0102_E0001, "設備コード")
            '    Return False
            'End If
            ''文字数は3文字以内
            'If dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, 1).Text.Length > 3 Then
            '    puErrMsg = String.Format(EXTM0102_E0010, "設備名", "3")
            '    Return False
            'End If

            '付帯設備名
            'For j = 1 To dataEXTM0102.PropVwFutaiSheet.ActiveSheet.ColumnCount - 1
            '    If dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, 1).Value = Nothing And dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, j).Value <> Nothing Then
            '        errCnt += 1
            '        Exit For
            '    End If
            'Next
            If dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, M0102_FUTAI_COL_SETUBINM).Value = Nothing Then
                For j = 1 To dataEXTM0102.PropVwFutaiSheet.ActiveSheet.ColumnCount - 4
                    If dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, j).Value <> Nothing Then
                        errCnt += 1
                        Exit For
                    End If
                    'If dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, j).Value = Nothing And i = 0 Then
                    '    errCnt2 += 1
                    '    Exit For
                    'End If
                Next
            End If
            'エラーがあればfalse返す
            If errCnt <> 0 Or errCnt2 <> 0 Then
                puErrMsg = String.Format(EXTM0102_E0001, "設備名")
                Return False
            End If
            ' 2015.12.15 UPD START↓ h.hagiwara
            '文字数は20文字以内
            'If dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, 1).Text.Length > 20 Then
            '    puErrMsg = String.Format(EXTM0102_E0010, "設備名", "20")
            '    Return False
            'End If
            If dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, M0102_FUTAI_COL_SETUBINM).Text.Length > 32 Then
                puErrMsg = String.Format(EXTM0102_E0010, "設備名", "32")
                Return False
            End If
            ' 2015.12.15 UPD END↑ h.hagiwara

            '単価
            'For j = 1 To dataEXTM0102.PropVwFutaiSheet.ActiveSheet.ColumnCount - 1
            '    If dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, 2).Value = Nothing And dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, j).Value <> Nothing Then
            '        errCnt += 1
            '        Exit For
            '    End If
            'Next
            If dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, M0102_FUTAI_COL_TANKA).Value = Nothing Then
                For j = 1 To dataEXTM0102.PropVwFutaiSheet.ActiveSheet.ColumnCount - 4
                    If dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, j).Value <> Nothing Then
                        errCnt += 1
                        Exit For
                    End If
                Next
            End If
            'エラーがあればfalse返す
            If errCnt <> 0 Then
                puErrMsg = String.Format(EXTM0102_E0001, "単価")
                Return False
            End If

            If dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, M0102_FUTAI_COL_TANKA).Value <> Nothing Then
                If ChkZenkaku(dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, M0102_FUTAI_COL_TANKA).Value) = False Then
                    puErrMsg = String.Format(EXTM0102_E0003, "単価")
                    Return False
                End If
                '単価は半角数字のみ
                If Not IsNumeric(dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, M0102_FUTAI_COL_TANKA).Value) Then
                    puErrMsg = String.Format(EXTM0102_E0003, "単価")
                    Return False
                End If
            End If

            '単位
            'For j = 1 To dataEXTM0102.PropVwFutaiSheet.ActiveSheet.ColumnCount - 1
            '    If dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, 3).Value = Nothing And dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, j).Value <> Nothing Then
            '        errCnt += 1
            '        Exit For
            '    End If
            'Next
            If dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, M0102_FUTAI_COL_TANNI).Value = Nothing Then
                For j = 1 To dataEXTM0102.PropVwFutaiSheet.ActiveSheet.ColumnCount - 4
                    If dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, j).Value <> Nothing Then
                        errCnt += 1
                        Exit For
                    End If
                Next
            End If
            'エラーがあればfalse返す
            If errCnt <> 0 Then
                puErrMsg = String.Format(EXTM0102_E0001, "単位")
                Return False
            End If

            '並び順
            'For j = 1 To dataEXTM0102.PropVwFutaiSheet.ActiveSheet.ColumnCount - 1
            '    If dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, 5).Value = Nothing And dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, j).Value <> Nothing Then
            '        errCnt += 1
            '        Exit For
            '    End If
            'Next
            If dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, M0102_FUTAI_COL_SORT).Value = Nothing Then
                For j = 1 To dataEXTM0102.PropVwFutaiSheet.ActiveSheet.ColumnCount - 4
                    If dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, j).Value <> Nothing Then
                        errCnt += 1
                        Exit For
                    End If
                Next
            End If
            'エラーがあればfalse返す
            If errCnt <> 0 Then
                puErrMsg = String.Format(EXTM0102_E0001, "並び順")
                Return False
            End If

            If dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, M0102_FUTAI_COL_SORT).Value <> Nothing Then
                If ChkZenkaku(dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, M0102_FUTAI_COL_SORT).Value) = False Then
                    puErrMsg = String.Format(EXTM0102_E0003, "並び順")
                    Return False
                End If

                '並び順は半角数字のみ
                If Not IsNumeric(dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, M0102_FUTAI_COL_SORT).Value) Then
                    puErrMsg = String.Format(EXTM0102_E0003, "並び順")
                    Return False
                End If
                '文字数は4文字以内
                If dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, M0102_FUTAI_COL_SORT).Text.Length > 4 Then
                    puErrMsg = String.Format(EXTM0102_E0010, "並び順", "4")
                    Return False
                End If
            End If
            If dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, M0102_FUTAI_COL_UPKBN).Value = "" Then
                dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, M0102_FUTAI_COL_UPKBN).Value = "1"
            End If
        Next

        '終了ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

        Return True

    End Function

    Private Function ChkZenkaku(ByVal ChkItem As String) As Boolean

        Dim r As New System.Text.RegularExpressions.Regex("^[a-zA-Z0-9]+$")

        If r.IsMatch(ChkItem) = False Then
            Return False
        End If

        Return True

    End Function

    ''' <summary>
    ''' 編集済みコンボボックス取得処理
    ''' <paramref name="dataEXTM0102">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean  取得合否　True:取得成功 Flase:取得失敗</returns>
    ''' <remarks> 分類マスタより、データを取得
    ''' <para>作成情報：2015/08/13 yu.satoh
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function GetBunruicd(ByRef dataEXTM0102 As DataEXTM0102) As Boolean

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn1 As New NpgsqlConnection(DbString)    'コネクション
        Dim Adapter As New NpgsqlDataAdapter        'アダプタ
        Dim Table As New DataTable()                'テーブル

        Try

            'コネクションを開く
            Cn1.Open()

            '対象期間コンボボックス作成
            If sqlEXTM0102.SelectBunruiCD(Adapter, Cn1, dataEXTM0102) = False Then
                Return False
            End If

            'SQLログ
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "システム管理情報", Nothing, Adapter.SelectCommand)

            'データを取得
            Adapter.Fill(Table)
            '取得データをデータクラスへ保存
            If IsDBNull(Table.Rows(0).Item(0)) Then
                dataEXTM0102.PropMaxBunruiCd = "01"
            Else
                dataEXTM0102.PropMaxBunruiCd = Format(Integer.Parse(Table.Rows(0).Item(0)) + 1, "00")
            End If

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            Return True
        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            puErrMsg = EXTM0102_E0000 & ex.Message
            Return False
        Finally
            Cn1.Close()
            Cn1.Dispose()
            Adapter.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' 編集済みコンボボックス取得処理
    ''' <paramref name="dataEXTM0102">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean  取得合否　True:取得成功 Flase:取得失敗</returns>
    ''' <remarks> 分類マスタより、データを取得
    ''' <para>作成情報：2015/08/13 yu.satoh
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function GetFutaiBunruicd(ByRef dataEXTM0102 As DataEXTM0102) As Boolean

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn1 As New NpgsqlConnection(DbString)    'コネクション
        Dim Adapter As New NpgsqlDataAdapter        'アダプタ
        Dim Table As New DataTable()                'テーブル

        Try

            'コネクションを開く
            Cn1.Open()

            '対象期間コンボボックス作成
            If sqlEXTM0102.SelectFutaiBunruiCD(Adapter, Cn1, dataEXTM0102) = False Then
                Return False
            End If

            'SQLログ
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "システム管理情報", Nothing, Adapter.SelectCommand)

            'データを取得
            Adapter.Fill(Table)
            '取得データをデータクラスへ保存
            If IsDBNull(Table.Rows(0).Item(0)) Then
                dataEXTM0102.PropMaxFutaiBunruiCd = 0
            Else
                dataEXTM0102.PropMaxFutaiBunruiCd = Integer.Parse(Table.Rows(0).Item(0))
            End If

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            Return True
        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            puErrMsg = EXTM0102_E0000 & ex.Message
            Return False
        Finally
            Cn1.Close()
            Cn1.Dispose()
            Adapter.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' ＤＢ更新処理
    ''' <paramref name="dataEXTM0102">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean  取得合否　True:取得成功 Flase:取得失敗</returns>
    ''' <remarks> ＤＢ更新を行う
    ''' <para>作成情報：2015/08/13 yu.satoh
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function InsertUpdateDB(ByRef dataEXTM0102 As DataEXTM0102) As Boolean

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)  'サーバーとクライアントをつなげる
        Dim Tsx As NpgsqlTransaction = Nothing    'トランザクション
        Dim noData As Integer = 0       '分類
        Dim noDataFutai As Integer = 0  '付帯設備

        Try

            'コネクションを開く
            Cn.Open()

            'トランザクションレベルを設定し、トランザクションを開始する
            Tsx = Cn.BeginTransaction(IsolationLevel.Serializable)

            ' 最新値を取得する
            If GetBunruicd(dataEXTM0102) = False Then
                Return False
            End If

            For i = 0 To dataEXTM0102.PropVwGroupingSheet.ActiveSheet.RowCount - 1

                '変数を初期化
                noData = 0

                '空の行を判断
                For j = 0 To dataEXTM0102.PropVwGroupingSheet.ActiveSheet.ColumnCount - 1
                    If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, j).Value <> Nothing Then

                    Else
                        noData += 1
                    End If
                Next

                '更新、もしくは挿入実行、全て空の行では実行しない
                'If noData <> 15 Then

                If noData <> 11 Then
                    If EntryBunrui(i, dataEXTM0102, noData, Tsx, Cn) = False Then
                        MsgBox(puErrMsg)
                        Return False
                    End If
                    If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, M0102_BUNRUI_COL_BUNRUICD).Value = Nothing Then
                        dataEXTM0102.PropUpdBunruiCd = dataEXTM0102.PropMaxBunruiCd
                        dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, M0102_BUNRUI_COL_BUNRUICD).Value = dataEXTM0102.PropMaxBunruiCd
                        dataEXTM0102.PropMaxBunruiCd = Format(Integer.Parse(dataEXTM0102.PropMaxBunruiCd) + 1, "00")
                    Else
                        If dataEXTM0102.PropMaxBunruiCd < dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, M0102_BUNRUI_COL_BUNRUICD).Value Then
                            If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, M0102_BUNRUI_COL_BUNRUICD).Value <> "99" Then
                                dataEXTM0102.PropMaxBunruiCd = dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, M0102_BUNRUI_COL_BUNRUICD).Value
                                dataEXTM0102.PropMaxBunruiCd = Format(Integer.Parse(dataEXTM0102.PropMaxBunruiCd) + 1, "00")
                            End If
                        End If
                    End If
                    ''付帯設備表表示行数で、値が全て空欄のもの以外実行
                    'For x = 0 To dataEXTM0102.PropVwFutaiSheet.ActiveSheet.RowCount - 1
                    '    If dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(x, 7).Value = i Then
                    '        '変数を初期化
                    '        noDataFutai = 0
                    '        '値が
                    '        For y = 0 To dataEXTM0102.PropVwFutaiSheet.ActiveSheet.ColumnCount - 2
                    '            If dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(x, y).Value <> Nothing Then

                    '            Else
                    '                noDataFutai += 1
                    '            End If
                    '        Next
                    '        '更新、もしくは挿入実行、全て空の行では実行しない
                    '        If noDataFutai <> 7 Then

                    '            If EntryFutai(x, dataEXTM0102, noDataFutai, Tsx, Cn) = False Then
                    '                MsgBox(puErrMsg)
                    '                Return False
                    '            End If
                    '        End If
                    '    End If
                    'Next
                End If
            Next

            If EntryFutai(dataEXTM0102, Tsx, Cn) = False Then
                MsgBox(puErrMsg)
                Return False
            End If

            ' 2015.12.25 ADD START↓ h.hagiwara
            If dataEXTM0102.PropTheaterBtn.Checked = True Then
                If GetRiyoInf(dataEXTM0102, Cn) = False Then
                    If EntryRiyoInf(dataEXTM0102, Tsx, Cn) = False Then
                        MsgBox(puErrMsg)
                        Return False
                    End If
                End If
            End If
            ' 2015.12.25 ADD END↑ h.hagiwara

            'コミット
            Tsx.Commit()

            'コネクションを閉じる
            Cn.Close()

            Return True

        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            'ロールバック
            If Tsx IsNot Nothing Then
                Tsx.Rollback()
            End If
            'コネクションが閉じられていない場合は閉じる
            If Cn IsNot Nothing Then
                Cn.Close()
            End If
            puErrMsg = EXTM0102_E0000 & ex.Message
            Return False
        Finally
            If Tsx IsNot Nothing Then
                Tsx.Dispose()
            End If
            Cn.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' 料金データ再セット
    ''' </summary>
    ''' <returns>boolean  エラーコード    true  正常終了  false  異常終了</returns>
    ''' <remarks>SPREAD上に表示している情報と退避した情報をマージする
    ''' <para>作成情報：2015.10.23 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function RyokinInfSet(ByRef dataEXTM0102 As DataEXTM0102) As Boolean
        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Dim dtrow1 As DataRow

        Try

            dataEXTM0102.PropDtFutaiMst.Clear()

            For i = 0 To dataEXTM0102.PropDtFutaiMstEscp.Rows.Count - 1
                dtrow1 = dataEXTM0102.PropDtFutaiMst.NewRow
                For k = 0 To 9
                    dtrow1(k) = dataEXTM0102.PropDtFutaiMstEscp.Rows(i).Item(k)
                Next
                dataEXTM0102.PropDtFutaiMst.Rows.Add(dtrow1)
            Next

            ' 格納領域を表示用・退避用をマージする。
            For i = 0 To dataEXTM0102.PropVwFutaiSheet.Sheets(0).RowCount - 1
                If dataEXTM0102.PropVwFutaiSheet.Sheets(0).Cells(i, M0102_FUTAI_COL_SETUBINM).Text = "" Then
                Else
                    dtrow1 = dataEXTM0102.PropDtFutaiMst.NewRow
                    dtrow1(0) = dataEXTM0102.PropVwFutaiSheet.Sheets(0).Cells(i, M0102_FUTAI_COL_SETUBICD).Value
                    dtrow1(1) = dataEXTM0102.PropVwFutaiSheet.Sheets(0).Cells(i, M0102_FUTAI_COL_SETUBINM).Value
                    dtrow1(2) = dataEXTM0102.PropVwFutaiSheet.Sheets(0).Cells(i, M0102_FUTAI_COL_TANKA).Value
                    dtrow1(3) = dataEXTM0102.PropVwFutaiSheet.Sheets(0).Cells(i, M0102_FUTAI_COL_TANNI).Value
                    If dataEXTM0102.PropVwFutaiSheet.Sheets(0).Cells(i, M0102_FUTAI_COL_DEFFLG).Value = True Then
                        dtrow1(4) = "1"
                    Else
                        dtrow1(4) = "0"
                    End If
                    dtrow1(5) = dataEXTM0102.PropVwFutaiSheet.Sheets(0).Cells(i, M0102_FUTAI_COL_SORT).Value
                    If dataEXTM0102.PropVwFutaiSheet.Sheets(0).Cells(i, M0102_FUTAI_COL_STS).Value = True Then
                        dtrow1(6) = "1"
                    Else
                        dtrow1(6) = "0"
                    End If
                    dtrow1(7) = dataEXTM0102.PropVwFutaiSheet.Sheets(0).Cells(i, M0102_FUTAI_COL_ROW).Value
                    dtrow1(8) = dataEXTM0102.PropVwFutaiSheet.Sheets(0).Cells(i, M0102_FUTAI_COL_BUNRUICD).Value
                    dtrow1(9) = dataEXTM0102.PropVwFutaiSheet.Sheets(0).Cells(i, M0102_FUTAI_COL_UPKBN).Value
                    dataEXTM0102.PropDtFutaiMst.Rows.Add(dtrow1)
                End If
            Next

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常処理終了正常処理終了
            Return True
        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            '例外処理
            puErrMsg = E0000 & ex.Message

            Return False
        End Try

    End Function

    ''' <summary>
    ''' 料金データ再セット
    ''' </summary>
    ''' <returns>boolean  エラーコード    true  正常終了  false  異常終了</returns>
    ''' <remarks>SPREAD上に表示している情報と退避した情報をマージする
    ''' <para>作成情報：2015.10.23 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function InitFutaiBunruiSet(ByRef dataEXTM0102 As DataEXTM0102) As Boolean
        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try
            ' 格納領域を表示用・退避用をマージする。
            If dataEXTM0102.PropDtFbunruiMst.Rows.Count > 0 Then
                For i = 0 To dataEXTM0102.PropDtFutaiMst.Rows.Count - 1
                    For j = 0 To dataEXTM0102.PropDtFbunruiMst.Rows.Count - 1
                        If dataEXTM0102.PropDtFutaiMst.Rows(i).Item(8).ToString = dataEXTM0102.PropDtFbunruiMst.Rows(j).Item(6).ToString Then
                            dataEXTM0102.PropDtFutaiMst.Rows(i).Item(7) = j
                        End If
                    Next
                Next
            End If

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常処理終了正常処理終了
            Return True
        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            '例外処理
            puErrMsg = E0000 & ex.Message

            Return False
        End Try

    End Function

    ''' <summary>
    ''' シアター利用料用データ登録処理
    ''' <paramref name="dataEXTM0102">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <remarks>登録ボタン押下時、付帯設備表の項目をインサート、もしくはアップデート処理
    ''' <para>作成情報：2015.12.25 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function EntryRiyoInf(dataEXTM0102 As DataEXTM0102 _
                               , ByRef Tsx As NpgsqlTransaction, _
                                          ByVal Cn As NpgsqlConnection) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Adapter As New NpgsqlCommand        'アダプタ
        Dim Table As New DataTable()                'テーブル

        Try
            ' 2016.01.05 DEL START↓ h.hagiwara
            'If GetCopyRiyoryoinf(dataEXTM0102) = True Then
            '    'SQLログ
            '    If sqlEXTM0102.InsertCopyRiyoInf(Adapter, Tsx, Cn, dataEXTM0102) = False Then
            '        Return False
            '    End If
            '    'SQLログ
            '    If sqlEXTM0102.InsertCopyRiyoFutaiInf(Adapter, Tsx, Cn, dataEXTM0102) = False Then
            '        Return False
            '    End If
            'Else
            ' 2016.01.05 DEL END↑ h.hagiwara
            ' 科目コードなどを取得
            If GeKamokuinf(dataEXTM0102) = True Then
                'SQLログ
                If sqlEXTM0102.InsertRiyoryoInf(Adapter, Tsx, Cn, dataEXTM0102) = False Then
                    Return False
                End If
            Else
                Return False
            End If
            'End If                                      ' 2016.01.05 DEL h.hagiwara

            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "システム管理情報", Nothing, Adapter)

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            Return True

        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            'ロールバック
            If Tsx IsNot Nothing Then
                Tsx.Rollback()
            End If
            puErrMsg = EXTM0102_E0000 & ex.Message
            Return False
        Finally
            'Cn.Close()
            'Cn.Dispose()
            Adapter.Dispose()
            Table.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' シアター利用料用データ存在チェック処理
    ''' <paramref name="dataEXTM0102">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean  取得合否　True:取得成功 Flase:取得失敗</returns>
    ''' <remarks> 登録・更新内容にシアター利用料用データが存在するかチェックする
    ''' <para>作成情報：2015.12.25 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function GetRiyoInf(ByRef dataEXTM0102 As DataEXTM0102, ByVal Cn As NpgsqlConnection) As Boolean

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言   
        Dim noData As Integer = 0                   '分類
        Dim Adapter As New NpgsqlDataAdapter        'アダプタ
        Dim Table As New DataTable                  'テーブル

        Try
            For i = 0 To dataEXTM0102.PropVwGroupingSheet.ActiveSheet.RowCount - 1

                '変数を初期化
                noData = 0

                '空の行を判断
                For j = 0 To dataEXTM0102.PropVwGroupingSheet.ActiveSheet.ColumnCount - 1
                    If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, j).Value <> Nothing Then

                    Else
                        noData += 1
                    End If
                Next

                If noData <> 15 Then
                    ' 登録・更新対象で科目マスタの利用料科目フラグ＝”１”のデータが存在するか確認
                    If sqlEXTM0102.GetRiyoKamokuInf(Adapter, Cn, dataEXTM0102, i) = False Then
                        Return False
                    End If

                    'データを取得
                    Table.Clear()
                    Adapter.Fill(Table)
                    '取得データをデータクラスへ保存
                    If Table.Rows.Count > 0 Then
                        If IsDBNull(Table.Rows(0).Item(0)) = False Then
                            If Table.Rows(0).Item(0) = "1" Then
                                Return True
                            End If
                        End If
                    End If
                End If
            Next

            Return False

        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            puErrMsg = EXTM0102_E0000 & ex.Message
            Return False
        Finally
            Adapter.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' シアター利用料用科目情報取得処理
    ''' <paramref name="dataEXTM0102">[IN/OUT]データクラス</paramref>
    ''' </summary>
    ''' <returns>boolean  取得合否　True:取得成功 Flase:取得失敗</returns>
    ''' <remarks> 登録済情報にシアター利用料用データが存在するかチェックする
    ''' <para>作成情報：2015.12.25 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function GeKamokuinf(ByRef dataEXTM0102 As DataEXTM0102) As Boolean

        '開始ログ出力()
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言   
        Dim Cn1 As New NpgsqlConnection(DbString)    ' コネクション
        Dim Adapter As New NpgsqlDataAdapter         ' アダプタ
        Dim Table As New DataTable()                 ' テーブル

        Try
            'コネクションを開く
            Cn1.Open()

            ' 登録済で科目マスタの利用料科目フラグ＝”１”のデータが存在するか確認
            If sqlEXTM0102.SqlKamokuinf(Adapter, Cn1, dataEXTM0102) = False Then
                Return False
            End If

            'データを取得
            Adapter.Fill(Table)
            '取得データをデータクラスへ保存
            If Table.Rows.Count > 0 Then
                dataEXTM0102.PropDtCopyKamokuMst = Table
            Else
                Return False
            End If

            Return True

        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            puErrMsg = EXTM0102_E0000 & ex.Message
            Return False
        Finally
            Adapter.Dispose()
        End Try

    End Function

End Class
