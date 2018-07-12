Imports Common
Imports CommonEXT
Imports Npgsql

''' <summary>
''' LogicEXTZ0103
''' </summary>
''' <remarks>日別請求一覧画面で発生する情報の取得処理等を行う
''' <para>作成情報：2015/09/16 h.endo
''' <p>改訂情報:</p>
''' </para></remarks>
Public Class LogicEXTZ0103

    '変数宣言
    Private sqlEXTZ0103 As New SqlEXTZ0103              'sqlクラス

#Region "スプレッドの作成"

    ''' <summary>
    ''' スプレッドの作成
    ''' </summary>
    ''' <param name="dataEXTZ0103">データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了　false 異常終了</returns>
    ''' <remarks>スプレッドを作成する。
    ''' <para>作成情報：2015/09/17 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Function MakeSpread(ByRef dataEXTZ0103 As DataEXTZ0103) As Boolean

        '変数宣言
        Dim intlastRow As Integer = 0               '各一覧最終行
        ' 2016.02.15 UPD START↓ h.hagiwara
        'Dim intRiyoKin As Integer = 0              '利用料
        'Dim intFutaiKin As Integer = 0             '付帯設備使用料
        'Dim intFutaiChosei As Integer = 0          '付帯設備調整額
        'Dim intKihonUriageGokei As Integer = 0     '基本売上合計
        Dim intRiyoKin As Long = 0                  '利用料
        Dim intFutaiKin As Long = 0                 '付帯設備使用料
        Dim intFutaiChosei As Long = 0              '付帯設備調整額
        Dim intKihonUriageGokei As Long = 0         '基本売上合計
        ' 2016.02.15 UPD END↑ h.hagiwara
        Dim intDrinkGenkin As Integer = 0           'ドリンク現金
        Dim intOneDrink As Integer = 0              'ワンドリンク 
        Dim intCoinLocker As Integer = 0            'コインロッカー
        Dim intAki1 As Integer = 0                  '空き1
        Dim intAki2 As Integer = 0                  '空き2
        Dim intAki3 As Integer = 0                  '空き3
        Dim intAki4 As Integer = 0                  '空き4
        Dim intZatsuShunyu As Integer = 0           '雑収入
        Dim intSuikaGenkin As Integer = 0           'SUIKA現金
        Dim intSuikaCoinLocker As Integer = 0       'SUIKAコインロッカー
        Dim intSonota As Integer = 0                'その他
        Dim intGenkinGokei As Integer = 0           '現金合計
        ' 2016.02.15 UPD START↓ h.hagiwara
        'Dim intSoUriageGokei As Integer = 0        '総売上合計
        'Dim intChoseiKin As Integer = 0            '調整額
        Dim intSoUriageGokei As Long = 0            '総売上合計
        Dim intChoseiKin As Long = 0                '調整額
        ' 2016.02.15 UPD END↑ h.hagiwara

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try

            With dataEXTZ0103

                'データ取得
                If GetSpreadData(dataEXTZ0103) = False Then
                    Return False
                End If

                '一覧表示
                If .PropStrShisetsuKbn = SHISETU_KBN_THEATER Then
                    'シアターの場合

                    '---一覧クリア
                    If .PropvwDayUriageTheatre.ActiveSheet.Rows.Count > 0 Then
                        .PropvwDayUriageTheatre.Sheets(0).RemoveRows(0, .PropvwDayUriageTheatre.ActiveSheet.Rows.Count)
                    End If
                    If .PropvwSeikyuChoseiTheatre.ActiveSheet.Rows.Count > 0 Then
                        .PropvwSeikyuChoseiTheatre.Sheets(0).RemoveRows(0, .PropvwSeikyuChoseiTheatre.ActiveSheet.Rows.Count)
                    End If

                    '---日別売上一覧表示
                    intlastRow = .PropvwDayUriageTheatre.ActiveSheet.RowCount
                    For intDataRow As Integer = 0 To .PropDtDayUriage.Rows.Count - 1

                        '最終行に1行追加
                        .PropvwDayUriageTheatre.Sheets(0).AddUnboundRows(intlastRow, 1)

                        '予約番号
                        .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_YoyakuNo).Value = .PropDtDayUriage.Rows(intDataRow)("YOYAKU_NO").ToString
                        '利用日
                        .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_RiyoDt).Value = .PropDtDayUriage.Rows(intDataRow)("YOYAKU_DT").ToString
                        '催事名
                        .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_SaijiNm).Value = .PropDtDayUriage.Rows(intDataRow)("SAIJI_NM").ToString
                        '利用者名
                        .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_RiyoNm).Value = .PropDtDayUriage.Rows(intDataRow)("RIYO_NM").ToString
                        '貸出種別
                        .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_KashiKind).Value = .PropDtDayUriage.Rows(intDataRow)("KASHI_KIND").ToString
                        '利用形状
                        .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_RiyoType).Value = .PropDtDayUriage.Rows(intDataRow)("RIYO_TYPE").ToString
                        '催事分類
                        .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_SaijiBunrui).Value = .PropDtDayUriage.Rows(intDataRow)("SAIJI_BUNRUI").ToString
                        '利用料
                        intRiyoKin = .PropDtDayUriage.Rows(intDataRow)("RIYO_KIN")
                        .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_RiyoKin).Value = intRiyoKin.ToString("#,##0")
                        '付帯設備使用料
                        intFutaiKin = .PropDtDayUriage.Rows(intDataRow)("FUTAI_KIN")
                        .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_FutaiKin).Value = intFutaiKin.ToString("#,##0")
                        '付帯設備調整額
                        intFutaiChosei = .PropDtDayUriage.Rows(intDataRow)("FUTAI_CHOSEI")
                        .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_FutaiChosei).Value = intFutaiChosei.ToString("#,##0")
                        '基本売上合計
                        intKihonUriageGokei = intRiyoKin + intFutaiKin + intFutaiChosei
                        .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_KihonUriageGokei).Value = intKihonUriageGokei.ToString("#,##0")
                        If intKihonUriageGokei < 0 Then
                            .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_KihonUriageGokei).ForeColor = Color.Red
                        Else
                            .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_KihonUriageGokei).ForeColor = Color.Black
                        End If
                        'ドリンク現金
                        intDrinkGenkin = .PropDtDayUriage.Rows(intDataRow)("DRINK_GENKIN")
                        .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_DrinkGenkin).Value = intDrinkGenkin.ToString("#,##0")
                        'ワンドリンク
                        intOneDrink = .PropDtDayUriage.Rows(intDataRow)("ONE_DRINK")
                        .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_OneDrink).Value = intOneDrink.ToString("#,##0")
                        'コインロッカー
                        intCoinLocker = .PropDtDayUriage.Rows(intDataRow)("COIN_LOCKER")
                        .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_CoinLocker).Value = intCoinLocker.ToString("#,##0")
                        '雑収入
                        intZatsuShunyu = .PropDtDayUriage.Rows(intDataRow)("ZATSU_SHUNYU")
                        .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_ZatsuShunyu).Value = intZatsuShunyu.ToString("#,##0")
                        '空き1
                        intAki1 = .PropDtDayUriage.Rows(intDataRow)("AKI1")
                        .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_Aki1).Value = intAki1.ToString("#,##0")
                        '空き2
                        intAki2 = .PropDtDayUriage.Rows(intDataRow)("AKI2")
                        .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_Aki2).Value = intAki2.ToString("#,##0")
                        '空き3
                        intAki3 = .PropDtDayUriage.Rows(intDataRow)("AKI3")
                        .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_Aki3).Value = intAki3.ToString("#,##0")
                        '空き4
                        intAki4 = .PropDtDayUriage.Rows(intDataRow)("AKI4")
                        .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_Aki4).Value = intAki4.ToString("#,##0")
                        'SUIKA現金
                        intSuikaGenkin = .PropDtDayUriage.Rows(intDataRow)("SUICA_GENKIN")
                        .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_SuikaGenkin).Value = intSuikaGenkin.ToString("#,##0")
                        'SUIKAコインロッカー
                        intSuikaCoinLocker = .PropDtDayUriage.Rows(intDataRow)("SUICA_COIN_LOCKER")
                        .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_SuikaCoinLocker).Value = intSuikaCoinLocker.ToString("#,##0")
                        'その他
                        intSonota = .PropDtDayUriage.Rows(intDataRow)("SONOTA")
                        .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_Sonota).Value = intSonota.ToString("#,##0")
                        '現金合計
                        intGenkinGokei = intDrinkGenkin + intOneDrink + intCoinLocker + intZatsuShunyu _
                                        + intAki1 + intAki2 + intAki3 + intAki4 _
                                        + intSuikaGenkin + intSuikaCoinLocker + intSonota
                        .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_GenkinGokei).Value = intGenkinGokei.ToString("#,##0")
                        '総売上合計
                        intSoUriageGokei = intKihonUriageGokei + intGenkinGokei
                        .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_SoUriageGokei).Value = intSoUriageGokei.ToString("#,##0")
                        If intSoUriageGokei < 0 Then
                            .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_SoUriageGokei).ForeColor = Color.Red
                        Else
                            .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_SoUriageGokei).ForeColor = Color.Black
                        End If

                        intlastRow += 1
                    Next

                    '---請求時の調整額表示
                    intlastRow = .PropvwSeikyuChoseiTheatre.ActiveSheet.RowCount
                    For intDataRow As Integer = 0 To .PropDtSeikyuChosei.Rows.Count - 1

                        '最終行に1行追加
                        .PropvwSeikyuChoseiTheatre.Sheets(0).AddUnboundRows(intlastRow, 1)

                        '予約番号
                        .PropvwSeikyuChoseiTheatre.ActiveSheet.Cells(intlastRow, SpreadSeikyuChoseiIndex_Theatre_YoyakuNo).Value = .PropDtSeikyuChosei.Rows(intDataRow)("YOYAKU_NO").ToString
                        '利用日
                        .PropvwSeikyuChoseiTheatre.ActiveSheet.Cells(intlastRow, SpreadSeikyuChoseiIndex_Theatre_RiyoDt).Value = .PropDtSeikyuChosei.Rows(intDataRow)("YOYAKU_DT").ToString
                        '催事名
                        .PropvwSeikyuChoseiTheatre.ActiveSheet.Cells(intlastRow, SpreadSeikyuChoseiIndex_Theatre_SaijiNm).Value = .PropDtSeikyuChosei.Rows(intDataRow)("SAIJI_NM").ToString
                        '利用者名
                        .PropvwSeikyuChoseiTheatre.ActiveSheet.Cells(intlastRow, SpreadSeikyuChoseiIndex_Theatre_RiyoNm).Value = .PropDtSeikyuChosei.Rows(intDataRow)("RIYO_NM").ToString
                        '貸出種別
                        .PropvwSeikyuChoseiTheatre.ActiveSheet.Cells(intlastRow, SpreadSeikyuChoseiIndex_Theatre_KashiKind).Value = .PropDtSeikyuChosei.Rows(intDataRow)("KASHI_KIND").ToString
                        '利用形状
                        .PropvwSeikyuChoseiTheatre.ActiveSheet.Cells(intlastRow, SpreadSeikyuChoseiIndex_Theatre_RiyoType).Value = .PropDtSeikyuChosei.Rows(intDataRow)("RIYO_TYPE").ToString
                        '催事分類
                        .PropvwSeikyuChoseiTheatre.ActiveSheet.Cells(intlastRow, SpreadSeikyuChoseiIndex_Theatre_SaijiBunrui).Value = .PropDtSeikyuChosei.Rows(intDataRow)("SAIJI_BUNRUI").ToString
                        '請求依頼番号
                        .PropvwSeikyuChoseiTheatre.ActiveSheet.Cells(intlastRow, SpreadSeikyuChoseiIndex_Theatre_SeikyuIraiNo).Value = .PropDtSeikyuChosei.Rows(intDataRow)("SEIKYU_IRAI_NO").ToString
                        '請求内容
                        .PropvwSeikyuChoseiTheatre.ActiveSheet.Cells(intlastRow, SpreadSeikyuChoseiIndex_Theatre_SeikyuNaiyo).Value = .PropDtSeikyuChosei.Rows(intDataRow)("SEIKYU_NAIYO").ToString
                        '調整額
                        intChoseiKin = .PropDtSeikyuChosei.Rows(intDataRow)("CHOSEI_KIN")
                        .PropvwSeikyuChoseiTheatre.ActiveSheet.Cells(intlastRow, SpreadSeikyuChoseiIndex_Theatre_ChoseiKin).Value = intChoseiKin.ToString("#,##0")
                        If intChoseiKin < 0 Then
                            .PropvwSeikyuChoseiTheatre.ActiveSheet.Cells(intlastRow, SpreadSeikyuChoseiIndex_Theatre_ChoseiKin).ForeColor = Color.Red
                        Else
                            .PropvwSeikyuChoseiTheatre.ActiveSheet.Cells(intlastRow, SpreadSeikyuChoseiIndex_Theatre_ChoseiKin).ForeColor = Color.Black
                        End If

                        intlastRow += 1
                    Next
                Else
                    'スタジオの場合

                    '---一覧クリア
                    If .PropvwDayUriageStudio.ActiveSheet.Rows.Count > 0 Then
                        .PropvwDayUriageStudio.Sheets(0).RemoveRows(0, .PropvwDayUriageStudio.ActiveSheet.Rows.Count)
                    End If
                    If .PropvwSeikyuChoseiStudio.ActiveSheet.Rows.Count > 0 Then
                        .PropvwSeikyuChoseiStudio.Sheets(0).RemoveRows(0, .PropvwSeikyuChoseiStudio.ActiveSheet.Rows.Count)
                    End If

                    '---日別売上一覧表示
                    intlastRow = .PropvwDayUriageStudio.ActiveSheet.RowCount
                    For intDataRow As Integer = 0 To .PropDtDayUriage.Rows.Count - 1

                        '最終行に1行追加
                        .PropvwDayUriageStudio.Sheets(0).AddUnboundRows(intlastRow, 1)

                        '予約番号
                        .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_YoyakuNo).Value = .PropDtDayUriage.Rows(intDataRow)("YOYAKU_NO").ToString
                        '利用日
                        .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_RiyoDt).Value = .PropDtDayUriage.Rows(intDataRow)("YOYAKU_DT").ToString
                        'アーティスト名
                        .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_ShutsuenNm).Value = .PropDtDayUriage.Rows(intDataRow)("SHUTSUEN_NM").ToString
                        '利用者名
                        .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_RiyoNm).Value = .PropDtDayUriage.Rows(intDataRow)("RIYO_NM").ToString
                        '貸出種別
                        .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_KashiKind).Value = .PropDtDayUriage.Rows(intDataRow)("KASHI_KIND").ToString
                        'スタジオ
                        .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_Studio).Value = .PropDtDayUriage.Rows(intDataRow)("STUDIO_KBN").ToString
                        '利用料
                        intRiyoKin = .PropDtDayUriage.Rows(intDataRow)("RIYO_KIN")
                        .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_RiyoKin).Value = intRiyoKin.ToString("#,##0")
                        '付帯設備使用料
                        intFutaiKin = .PropDtDayUriage.Rows(intDataRow)("FUTAI_KIN")
                        .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_FutaiKin).Value = intFutaiKin.ToString("#,##0")
                        '付帯設備調整額
                        intFutaiChosei = .PropDtDayUriage.Rows(intDataRow)("FUTAI_CHOSEI")
                        .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_FutaiChosei).Value = intFutaiChosei.ToString("#,##0")
                        '基本売上合計
                        intKihonUriageGokei = intRiyoKin + intFutaiKin + intFutaiChosei
                        .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_KihonUriageGokei).Value = intKihonUriageGokei.ToString("#,##0")
                        If intKihonUriageGokei < 0 Then
                            .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_KihonUriageGokei).ForeColor = Color.Red
                        Else
                            .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_KihonUriageGokei).ForeColor = Color.Black
                        End If
                        'ドリンク現金
                        intDrinkGenkin = .PropDtDayUriage.Rows(intDataRow)("DRINK_GENKIN")
                        .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_DrinkGenkin).Value = intDrinkGenkin.ToString("#,##0")
                        'ワンドリンク
                        intOneDrink = .PropDtDayUriage.Rows(intDataRow)("ONE_DRINK")
                        .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_OneDrink).Value = intOneDrink.ToString("#,##0")
                        'コインロッカー
                        intCoinLocker = .PropDtDayUriage.Rows(intDataRow)("COIN_LOCKER")
                        .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_CoinLocker).Value = intCoinLocker.ToString("#,##0")
                        '雑収入
                        intZatsuShunyu = .PropDtDayUriage.Rows(intDataRow)("ZATSU_SHUNYU")
                        .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_ZatsuShunyu).Value = intZatsuShunyu.ToString("#,##0")
                        '空き1
                        intAki1 = .PropDtDayUriage.Rows(intDataRow)("AKI1")
                        .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_Aki1).Value = intAki1.ToString("#,##0")
                        '空き2
                        intAki2 = .PropDtDayUriage.Rows(intDataRow)("AKI2")
                        .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_Aki2).Value = intAki2.ToString("#,##0")
                        '空き3
                        intAki3 = .PropDtDayUriage.Rows(intDataRow)("AKI3")
                        .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_Aki3).Value = intAki3.ToString("#,##0")
                        '空き4
                        intAki4 = .PropDtDayUriage.Rows(intDataRow)("AKI4")
                        .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_Aki4).Value = intAki4.ToString("#,##0")
                        'SUIKA現金
                        intSuikaGenkin = .PropDtDayUriage.Rows(intDataRow)("SUICA_GENKIN")
                        .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_SuikaGenkin).Value = intSuikaGenkin.ToString("#,##0")
                        'SUIKAコインロッカー
                        intSuikaCoinLocker = .PropDtDayUriage.Rows(intDataRow)("SUICA_COIN_LOCKER")
                        .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_SuikaCoinLocker).Value = intSuikaCoinLocker.ToString("#,##0")
                        'その他
                        intSonota = .PropDtDayUriage.Rows(intDataRow)("SONOTA")
                        .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_Sonota).Value = intSonota.ToString("#,##0")
                        '現金合計
                        intGenkinGokei = intDrinkGenkin + intOneDrink + intCoinLocker + intZatsuShunyu _
                                        + intAki1 + intAki2 + intAki3 + intAki4 _
                                        + intSuikaGenkin + intSuikaCoinLocker + intSonota
                        .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_GenkinGokei).Value = intGenkinGokei.ToString("#,##0")
                        '総売上合計
                        intSoUriageGokei = intKihonUriageGokei + intGenkinGokei
                        .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_SoUriageGokei).Value = intSoUriageGokei.ToString("#,##0")
                        If intSoUriageGokei < 0 Then
                            .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_SoUriageGokei).ForeColor = Color.Red
                        Else
                            .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_SoUriageGokei).ForeColor = Color.Black
                        End If

                        intlastRow += 1
                    Next

                     '---請求時の調整額表示
                    intlastRow = .PropvwSeikyuChoseiStudio.ActiveSheet.RowCount
                    For intDataRow As Integer = 0 To .PropDtSeikyuChosei.Rows.Count - 1

                        '最終行に1行追加
                        .PropvwSeikyuChoseiStudio.Sheets(0).AddUnboundRows(intlastRow, 1)

                        '予約番号
                        .PropvwSeikyuChoseiStudio.ActiveSheet.Cells(intlastRow, SpreadSeikyuChoseiIndex_Studio_YoyakuNo).Value = .PropDtSeikyuChosei.Rows(intDataRow)("YOYAKU_NO").ToString
                        '利用日
                        .PropvwSeikyuChoseiStudio.ActiveSheet.Cells(intlastRow, SpreadSeikyuChoseiIndex_Studio_RiyoDt).Value = .PropDtSeikyuChosei.Rows(intDataRow)("YOYAKU_DT").ToString
                        '催事名
                        .PropvwSeikyuChoseiStudio.ActiveSheet.Cells(intlastRow, SpreadSeikyuChoseiIndex_Studio_ShutsuenNm).Value = .PropDtSeikyuChosei.Rows(intDataRow)("SHUTSUEN_NM").ToString
                        '利用者名
                        .PropvwSeikyuChoseiStudio.ActiveSheet.Cells(intlastRow, SpreadSeikyuChoseiIndex_Studio_RiyoNm).Value = .PropDtSeikyuChosei.Rows(intDataRow)("RIYO_NM").ToString
                        '貸出種別
                        .PropvwSeikyuChoseiStudio.ActiveSheet.Cells(intlastRow, SpreadSeikyuChoseiIndex_Studio_KashiKind).Value = .PropDtSeikyuChosei.Rows(intDataRow)("KASHI_KIND").ToString
                        'スタジオ
                        .PropvwSeikyuChoseiStudio.ActiveSheet.Cells(intlastRow, SpreadSeikyuChoseiIndex_Studio_Studio).Value = .PropDtSeikyuChosei.Rows(intDataRow)("STUDIO_KBN").ToString
                        '請求依頼番号
                        .PropvwSeikyuChoseiStudio.ActiveSheet.Cells(intlastRow, SpreadSeikyuChoseiIndex_Studio_SeikyuIraiNo).Value = .PropDtSeikyuChosei.Rows(intDataRow)("SEIKYU_IRAI_NO").ToString
                        '請求内容
                        .PropvwSeikyuChoseiStudio.ActiveSheet.Cells(intlastRow, SpreadSeikyuChoseiIndex_Studio_SeikyuNaiyo).Value = .PropDtSeikyuChosei.Rows(intDataRow)("SEIKYU_NAIYO").ToString
                        '調整額
                        intChoseiKin = .PropDtSeikyuChosei.Rows(intDataRow)("CHOSEI_KIN")
                        .PropvwSeikyuChoseiStudio.ActiveSheet.Cells(intlastRow, SpreadSeikyuChoseiIndex_Studio_ChoseiKin).Value = intChoseiKin.ToString("#,##0")
                        If intChoseiKin < 0 Then
                            .PropvwSeikyuChoseiStudio.ActiveSheet.Cells(intlastRow, SpreadSeikyuChoseiIndex_Studio_ChoseiKin).ForeColor = Color.Red
                        Else
                            .PropvwSeikyuChoseiStudio.ActiveSheet.Cells(intlastRow, SpreadSeikyuChoseiIndex_Studio_ChoseiKin).ForeColor = Color.Black
                        End If

                        intlastRow += 1

                    Next

                End If

            End With

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

    '*****Privateメソッド*****
#Region "スプレッドへのデータ取得"

    ''' <summary>
    ''' スプレッドへのデータ取得
    ''' </summary>
    ''' <param name="dataEXTZ0103">データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了　false 異常終了</returns>
    ''' <remarks>スプレッドへのデータを取得する。
    ''' <para>作成情報：2015/09/17 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Private Function GetSpreadData(ByRef dataEXTZ0103 As DataEXTZ0103) As Boolean

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)        'コネクション
        Dim Adapter As New NpgsqlDataAdapter            'アダプタ

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try

            Cn.Open()

            '■日別売上データ取得
            'SELECT用SQLCommandを作成
            If sqlEXTZ0103.setSelectDayUriage(Adapter, Cn, dataEXTZ0103) = False Then
                '異常終了
                Return False
            End If

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "", Nothing, Adapter.SelectCommand)

            'データ取得
            dataEXTZ0103.PropDtDayUriage = New DataTable
            Adapter.Fill(dataEXTZ0103.PropDtDayUriage)

            '■請求調整データ取得
            'SELECT用SQLCommandを作成
            If sqlEXTZ0103.setSeikyuChosei(Adapter, Cn, dataEXTZ0103) = False Then
                '異常終了
                Return False
            End If

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "", Nothing, Adapter.SelectCommand)

            'データ取得
            dataEXTZ0103.PropDtSeikyuChosei = New DataTable
            Adapter.Fill(dataEXTZ0103.PropDtSeikyuChosei)

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
            If Cn IsNot Nothing Then
                Cn.Close()
            End If

            Adapter.Dispose()
            Cn.Dispose()

        End Try

    End Function

#End Region

End Class
