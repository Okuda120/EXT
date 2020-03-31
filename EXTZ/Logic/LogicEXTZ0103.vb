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

    ' --- 2019/08/23 日別売上一覧機能追加対応 Start E.Okuda@Compass ---
    Private Const STR_DELIMITER As String = "/"
    Private Const CSV_HEADER_THEATER As String = "予約番号,利用日,催事名,貸出種別,利用形状,催事分類," &
                                                 "利用料(A),付帯設備利用料(B),付帯設備調整額(C),基本売上合計(D)=(A+B+C)," &
                                                 "001ドリンク現金(E),002ワンドリンク(F),003コインロッカー(G),004雑収入(H)," &
                                                 "006空き(I),007空き(J),008空き(K),009空き(L),900SUICA現金(M),901SUICAコインロッカー(N)," &
                                                 "902-950その他(O),現金合計(P)=(E～O),総売上合計(Q)=(D+P)"
    Private Const CSV_HEADER_STUDIO As String = "予約番号,利用日,催事名,貸出種別,利用形状,催事分類," &
                                                 "利用料(A),付帯設備利用料(B),付帯設備調整額(C),基本売上合計(D)=(A+B+C)," &
                                                 "001ドリンク現金(E),002ワンドリンク(F),003コインロッカー(G),004雑収入(H)," &
                                                 "006空き(I),007空き(J),008空き(K),009空き(L),900SUICA現金(M),901SUICAコインロッカー(N)," &
                                                 "902-950その他(O),現金合計(P)=(E～O),総売上合計(Q)=(D+P)"

    ' --- 2019/08/23 日別売上一覧機能追加対応 End E.Okuda@Compass ---

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

        ' --- 2019/08/23 日別売上一覧機能追加対応 Start E.Okuda@Compass ---
        Dim strTempYoyakuNo As String = ""
        Dim strTempYoyakuDt As String = ""
        Dim strTempSaijiNm As String = ""
        Dim strTempShutsuenNm As String = ""
        Dim strTempRiyoNm As String = ""
        Dim strTempKashiKind As String = ""
        Dim strTempRiyoType As String = ""
        Dim strTempSaijiBunrui As String = ""
        Dim strTempStudioKbn As String = ""
        Dim lngTempRiyoKin As Long = 0
        Dim lngTempFutaiKin As Long = 0
        Dim lngTempFutaiChosei As Long = 0
        Dim lngTempDrinkGenkin As Long = 0
        Dim lngTempOneDrink As Long = 0
        Dim lngTempCoinLocker As Long = 0
        Dim lngTempZatsuShunyu As Long = 0
        Dim lngTempAki1 As Long = 0
        Dim lngTempAki2 As Long = 0
        Dim lngTempAki3 As Long = 0
        Dim lngTempAki4 As Long = 0
        Dim lngTempSuicaGenkin As Long = 0
        Dim lngTempSuicaCoinLocker As Long = 0
        Dim lngTempSonota As Long = 0
        Dim lngSumUriageGokei As Long = 0
        Dim lngSumGenkin As Long = 0
        Dim lngSumTotalUriage As Long = 0


        ' --- 2019/08/23 日別売上一覧機能追加対応 End E.Okuda@Compass ---


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

                    ' --- 2019/08/23 日別売上一覧機能追加対応 Start E.Okuda@Compass ---
                    If .PropDtDayUriage.Rows.Count > 0 Then
                        ' 書き込み用一時変数にデータを格納する。
                        strTempYoyakuNo = .PropDtDayUriage.Rows(0)("YOYAKU_NO").ToString
                        strTempYoyakuDt = .PropDtDayUriage.Rows(0)("YOYAKU_DT").ToString
                        strTempSaijiNm = .PropDtDayUriage.Rows(0)("SAIJI_NM").ToString
                        strTempRiyoNm = .PropDtDayUriage.Rows(0)("RIYO_NM").ToString
                        strTempKashiKind = .PropDtDayUriage.Rows(0)("KASHI_KIND").ToString()
                        strTempRiyoType = .PropDtDayUriage.Rows(0)("RIYO_TYPE").ToString()
                        strTempSaijiBunrui = .PropDtDayUriage.Rows(0)("SAIJI_BUNRUI").ToString()
                        lngTempRiyoKin = .PropDtDayUriage.Rows(0)("RIYO_KIN")
                        lngTempFutaiKin = .PropDtDayUriage.Rows(0)("FUTAI_SHOKEI")
                        lngTempFutaiChosei = .PropDtDayUriage.Rows(0)("FUTAI_CHOSEI")
                        lngTempDrinkGenkin = .PropDtDayUriage.Rows(0)("DRINK_GENKIN")
                        lngTempOneDrink = .PropDtDayUriage.Rows(0)("ONE_DRINK")
                        lngTempCoinLocker = .PropDtDayUriage.Rows(0)("COIN_LOCKER")
                        lngTempZatsuShunyu = .PropDtDayUriage.Rows(0)("ZATSU_SHUNYU")
                        lngTempAki1 = .PropDtDayUriage.Rows(0)("AKI1")
                        lngTempAki2 = .PropDtDayUriage.Rows(0)("AKI2")
                        lngTempAki3 = .PropDtDayUriage.Rows(0)("AKI3")
                        lngTempAki4 = .PropDtDayUriage.Rows(0)("AKI4")
                        lngTempSuicaGenkin = .PropDtDayUriage.Rows(0)("SUICA_GENKIN")
                        lngTempSuicaCoinLocker = .PropDtDayUriage.Rows(0)("SUICA_COIN_LOCKER")
                        lngTempSonota = .PropDtDayUriage.Rows(0)("SONOTA")
                        ' --- 2019/08/23 日別売上一覧機能追加対応 End E.Okuda@Compass ---

                        '---日別売上一覧表示
                        intlastRow = .PropvwDayUriageTheatre.ActiveSheet.RowCount
                        ' --- 2019/08/23 日別売上一覧機能追加対応 Start E.Okuda@Compass ---
                        For intDataRow As Integer = 1 To .PropDtDayUriage.Rows.Count - 1
                            'For intDataRow As Integer = 0 To .PropDtDayUriage.Rows.Count - 1
                            If (.PropDtDayUriage.Rows(intDataRow)("YOYAKU_DT").ToString <> strTempYoyakuDt Or
                                    .PropDtDayUriage.Rows(intDataRow)("SAIJI_NM").ToString <> strTempSaijiNm) Then
                                '最終行に1行追加
                                .PropvwDayUriageTheatre.Sheets(0).AddUnboundRows(intlastRow, 1)

                                ' 予約番号
                                .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_YoyakuNo).Value = strTempYoyakuNo
                                ' 利用日
                                .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_RiyoDt).Value = strTempYoyakuDt
                                '催事名
                                .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_SaijiNm).Value = strTempSaijiNm
                                '利用者名
                                .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_RiyoNm).Value = strTempRiyoNm
                                '貸出種別
                                .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_KashiKind).Value = strTempKashiKind
                                '利用形状
                                .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_RiyoType).Value = strTempRiyoType
                                '催事分類
                                .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_SaijiBunrui).Value = strTempSaijiBunrui
                                '利用料
                                .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_RiyoKin).Value = lngTempRiyoKin.ToString("#,##0")
                                '付帯設備使用料
                                .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_FutaiKin).Value = lngTempFutaiKin.ToString("#,##0")
                                '付帯設備調整額
                                .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_FutaiChosei).Value = lngTempFutaiChosei.ToString("#,##0")
                                '基本売上合計
                                lngSumUriageGokei = lngTempRiyoKin + lngTempFutaiKin + lngTempFutaiChosei
                                .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_KihonUriageGokei).Value = lngSumUriageGokei.ToString("#,##0")

                                If lngSumUriageGokei < 0 Then
                                    ' 合計額がマイナスなら赤文字にする。
                                    .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_KihonUriageGokei).ForeColor = Color.Red
                                Else
                                    .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_KihonUriageGokei).ForeColor = Color.Black
                                End If

                                'ドリンク現金
                                .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_DrinkGenkin).Value = lngTempDrinkGenkin.ToString("#,##0")
                                'ワンドリンク
                                .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_OneDrink).Value = lngTempOneDrink.ToString("#,##0")
                                'コインロッカー
                                .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_CoinLocker).Value = lngTempCoinLocker.ToString("#,##0")
                                '雑収入
                                .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_ZatsuShunyu).Value = lngTempZatsuShunyu.ToString("#,##0")
                                '空き1
                                .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_Aki1).Value = lngTempAki1.ToString("#,##0")
                                '空き2
                                .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_Aki2).Value = lngTempAki2.ToString("#,##0")
                                '空き3
                                .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_Aki3).Value = lngTempAki3.ToString("#,##0")
                                '空き4
                                .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_Aki4).Value = lngTempAki4.ToString("#,##0")
                                'SUIKA現金
                                .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_SuikaGenkin).Value = lngTempSuicaGenkin.ToString("#,##0")
                                'SUIKAコインロッカー
                                .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_SuikaCoinLocker).Value = lngTempSuicaCoinLocker.ToString("#,##0")
                                'その他
                                .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_Sonota).Value = lngTempSonota.ToString("#,##0")
                                '現金合計
                                lngSumGenkin = lngTempDrinkGenkin + lngTempOneDrink + lngTempCoinLocker + lngTempZatsuShunyu _
                                        + lngTempAki1 + lngTempAki2 + lngTempAki3 + lngTempAki4 _
                                        + lngTempSuicaGenkin + lngTempSuicaCoinLocker + lngTempSonota
                                .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_GenkinGokei).Value = lngSumGenkin.ToString("#,##0")
                                '総売上合計
                                lngSumTotalUriage = lngSumUriageGokei + lngSumGenkin
                                .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_SoUriageGokei).Value = lngSumTotalUriage.ToString("#,##0")
                                If lngSumTotalUriage < 0 Then
                                    ' 総売上がマイナスなら赤文字にする。
                                    .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_SoUriageGokei).ForeColor = Color.Red
                                Else
                                    .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_SoUriageGokei).ForeColor = Color.Black
                                End If

                                intlastRow += 1

                                ' 書き込み用一時変数にデータを格納する。
                                strTempYoyakuNo = .PropDtDayUriage.Rows(intDataRow)("YOYAKU_NO").ToString
                                strTempYoyakuDt = .PropDtDayUriage.Rows(intDataRow)("YOYAKU_DT").ToString
                                strTempSaijiNm = .PropDtDayUriage.Rows(intDataRow)("SAIJI_NM").ToString
                                strTempRiyoNm = .PropDtDayUriage.Rows(intDataRow)("RIYO_NM").ToString
                                strTempKashiKind = .PropDtDayUriage.Rows(intDataRow)("KASHI_KIND").ToString()
                                strTempRiyoType = .PropDtDayUriage.Rows(intDataRow)("RIYO_TYPE").ToString()
                                strTempSaijiBunrui = .PropDtDayUriage.Rows(intDataRow)("SAIJI_BUNRUI").ToString()
                                lngTempRiyoKin = .PropDtDayUriage.Rows(intDataRow)("RIYO_KIN")
                                lngTempFutaiKin = .PropDtDayUriage.Rows(intDataRow)("FUTAI_SHOKEI")
                                lngTempFutaiChosei = .PropDtDayUriage.Rows(intDataRow)("FUTAI_CHOSEI")
                                lngTempDrinkGenkin = .PropDtDayUriage.Rows(intDataRow)("DRINK_GENKIN")
                                lngTempOneDrink = .PropDtDayUriage.Rows(intDataRow)("ONE_DRINK")
                                lngTempCoinLocker = .PropDtDayUriage.Rows(intDataRow)("COIN_LOCKER")
                                lngTempZatsuShunyu = .PropDtDayUriage.Rows(intDataRow)("ZATSU_SHUNYU")
                                lngTempAki1 = .PropDtDayUriage.Rows(intDataRow)("AKI1")
                                lngTempAki2 = .PropDtDayUriage.Rows(intDataRow)("AKI2")
                                lngTempAki3 = .PropDtDayUriage.Rows(intDataRow)("AKI3")
                                lngTempAki4 = .PropDtDayUriage.Rows(intDataRow)("AKI4")
                                lngTempSuicaGenkin = .PropDtDayUriage.Rows(intDataRow)("SUICA_GENKIN")
                                lngTempSuicaCoinLocker = .PropDtDayUriage.Rows(intDataRow)("SUICA_COIN_LOCKER")
                                lngTempSonota = .PropDtDayUriage.Rows(intDataRow)("SONOTA")
                            Else
                                ' 予約番号
                                If Not strTempYoyakuNo = .PropDtDayUriage.Rows(intDataRow)("YOYAKU_NO").ToString Then
                                    strTempYoyakuNo = strTempYoyakuNo & STR_DELIMITER & .PropDtDayUriage.Rows(intDataRow)("YOYAKU_NO").ToString
                                End If
                                ' 利用日
                                If Not .PropDtDayUriage.Rows(intDataRow)("YOYAKU_DT").ToString = strTempYoyakuDt Then
                                    strTempYoyakuDt = strTempYoyakuDt & STR_DELIMITER & .PropDtDayUriage.Rows(intDataRow)("YOYAKU_DT").ToString
                                End If
                                '催事名
                                If Not .PropDtDayUriage.Rows(intDataRow)("SAIJI_NM").ToString = strTempSaijiNm Then
                                    strTempSaijiNm = strTempSaijiNm & STR_DELIMITER & .PropDtDayUriage.Rows(intDataRow)("SAIJI_NM").ToString()
                                End If
                                ' 利用者名
                                If Not strTempRiyoNm = .PropDtDayUriage.Rows(intDataRow)("RIYO_NM").ToString Then
                                    strTempRiyoNm = strTempRiyoNm & STR_DELIMITER & .PropDtDayUriage.Rows(intDataRow)("RIYO_NM").ToString
                                End If
                                ' 貸出種別
                                If Not strTempKashiKind = .PropDtDayUriage.Rows(intDataRow)("KASHI_KIND").ToString() Then
                                    strTempKashiKind = strTempKashiKind & STR_DELIMITER & .PropDtDayUriage.Rows(intDataRow)("KASHI_KIND").ToString()
                                End If
                                ' 利用形状
                                If Not strTempRiyoType = .PropDtDayUriage.Rows(intDataRow)("RIYO_TYPE").ToString() Then
                                    strTempRiyoType = strTempRiyoType & STR_DELIMITER & .PropDtDayUriage.Rows(intDataRow)("RIYO_TYPE").ToString()
                                End If
                                ' 催事分類
                                If Not strTempSaijiBunrui = .PropDtDayUriage.Rows(intDataRow)("SAIJI_BUNRUI").ToString() Then
                                    strTempSaijiBunrui = strTempSaijiBunrui & STR_DELIMITER & .PropDtDayUriage.Rows(intDataRow)("SAIJI_BUNRUI").ToString()
                                End If
                                ' 利用料
                                lngTempRiyoKin += .PropDtDayUriage.Rows(intDataRow)("RIYO_KIN")
                                ' 付帯設備使用料
                                lngTempFutaiKin += .PropDtDayUriage.Rows(intDataRow)("FUTAI_SHOKEI")
                                ' 付帯設備調整額
                                lngTempFutaiChosei += .PropDtDayUriage.Rows(intDataRow)("FUTAI_CHOSEI")
                                ' ドリンク現金
                                lngTempDrinkGenkin += .PropDtDayUriage.Rows(intDataRow)("DRINK_GENKIN")
                                ' ワンドリンク
                                lngTempOneDrink += .PropDtDayUriage.Rows(intDataRow)("ONE_DRINK")
                                ' コインロッカー
                                lngTempCoinLocker += .PropDtDayUriage.Rows(intDataRow)("COIN_LOCKER")
                                ' 雑収入
                                lngTempZatsuShunyu += .PropDtDayUriage.Rows(intDataRow)("ZATSU_SHUNYU")
                                ' 空き1
                                lngTempAki1 += .PropDtDayUriage.Rows(intDataRow)("AKI1")
                                ' 空き2
                                lngTempAki2 += .PropDtDayUriage.Rows(intDataRow)("AKI2")
                                ' 空き3
                                lngTempAki3 += .PropDtDayUriage.Rows(intDataRow)("AKI3")
                                ' 空き4
                                lngTempAki4 += .PropDtDayUriage.Rows(intDataRow)("AKI4")
                                ' SUIKA現金
                                lngTempSuicaGenkin += .PropDtDayUriage.Rows(intDataRow)("SUICA_GENKIN")
                                ' SUIKAコインロッカー
                                lngTempSuicaCoinLocker += .PropDtDayUriage.Rows(intDataRow)("SUICA_COIN_LOCKER")
                                ' その他
                                lngTempSonota = .PropDtDayUriage.Rows(intDataRow)("SONOTA")
                            End If

                            ''最終行に1行追加
                            '.PropvwDayUriageTheatre.Sheets(0).AddUnboundRows(intlastRow, 1)
                            ''予約番号
                            '.PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_YoyakuNo).Value = .PropDtDayUriage.Rows(intDataRow)("YOYAKU_NO").ToString
                            ''利用日
                            '.PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_RiyoDt).Value = .PropDtDayUriage.Rows(intDataRow)("YOYAKU_DT").ToString
                            ''催事名
                            '.PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_SaijiNm).Value = .PropDtDayUriage.Rows(intDataRow)("SAIJI_NM").ToString
                            ''利用者名
                            '.PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_RiyoNm).Value = .PropDtDayUriage.Rows(intDataRow)("RIYO_NM").ToString
                            ''貸出種別
                            '.PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_KashiKind).Value = .PropDtDayUriage.Rows(intDataRow)("KASHI_KIND").ToString
                            ''利用形状
                            '.PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_RiyoType).Value = .PropDtDayUriage.Rows(intDataRow)("RIYO_TYPE").ToString
                            ''催事分類
                            '.PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_SaijiBunrui).Value = .PropDtDayUriage.Rows(intDataRow)("SAIJI_BUNRUI").ToString
                            ''利用料
                            'intRiyoKin = .PropDtDayUriage.Rows(intDataRow)("RIYO_KIN")
                            '.PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_RiyoKin).Value = intRiyoKin.ToString("#,##0")
                            ''付帯設備使用料
                            'intFutaiKin = .PropDtDayUriage.Rows(intDataRow)("FUTAI_KIN")
                            '.PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_FutaiKin).Value = intFutaiKin.ToString("#,##0")
                            ''付帯設備調整額
                            'intFutaiChosei = .PropDtDayUriage.Rows(intDataRow)("FUTAI_CHOSEI")
                            '.PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_FutaiChosei).Value = intFutaiChosei.ToString("#,##0")
                            ''基本売上合計
                            'intKihonUriageGokei = intRiyoKin + intFutaiKin + intFutaiChosei
                            '.PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_KihonUriageGokei).Value = intKihonUriageGokei.ToString("#,##0")
                            'If intKihonUriageGokei <0 Then
                            '    .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_KihonUriageGokei).ForeColor = Color.Red
                            'Else
                            '    .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_KihonUriageGokei).ForeColor = Color.Black
                            'End If
                            ''ドリンク現金
                            'intDrinkGenkin = .PropDtDayUriage.Rows(intDataRow)("DRINK_GENKIN")
                            '.PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_DrinkGenkin).Value = intDrinkGenkin.ToString("#,##0")
                            ''ワンドリンク
                            'intOneDrink = .PropDtDayUriage.Rows(intDataRow)("ONE_DRINK")
                            '.PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_OneDrink).Value = intOneDrink.ToString("#,##0")
                            ''コインロッカー
                            'intCoinLocker = .PropDtDayUriage.Rows(intDataRow)("COIN_LOCKER")
                            '.PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_CoinLocker).Value = intCoinLocker.ToString("#,##0")
                            ''雑収入
                            'intZatsuShunyu = .PropDtDayUriage.Rows(intDataRow)("ZATSU_SHUNYU")
                            '.PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_ZatsuShunyu).Value = intZatsuShunyu.ToString("#,##0")
                            ''空き1
                            'intAki1 = .PropDtDayUriage.Rows(intDataRow)("AKI1")
                            '.PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_Aki1).Value = intAki1.ToString("#,##0")
                            ''空き2
                            'intAki2 = .PropDtDayUriage.Rows(intDataRow)("AKI2")
                            '.PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_Aki2).Value = intAki2.ToString("#,##0")
                            ''空き3
                            'intAki3 = .PropDtDayUriage.Rows(intDataRow)("AKI3")
                            '.PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_Aki3).Value = intAki3.ToString("#,##0")
                            ''空き4
                            'intAki4 = .PropDtDayUriage.Rows(intDataRow)("AKI4")
                            '.PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_Aki4).Value = intAki4.ToString("#,##0")
                            ''SUIKA現金
                            'intSuikaGenkin = .PropDtDayUriage.Rows(intDataRow)("SUICA_GENKIN")
                            '.PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_SuikaGenkin).Value = intSuikaGenkin.ToString("#,##0")
                            ''SUIKAコインロッカー
                            'intSuikaCoinLocker = .PropDtDayUriage.Rows(intDataRow)("SUICA_COIN_LOCKER")
                            '.PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_SuikaCoinLocker).Value = intSuikaCoinLocker.ToString("#,##0")
                            ''その他
                            'intSonota = .PropDtDayUriage.Rows(intDataRow)("SONOTA")
                            '.PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_Sonota).Value = intSonota.ToString("#,##0")
                            ''現金合計
                            'intGenkinGokei = intDrinkGenkin + intOneDrink + intCoinLocker + intZatsuShunyu _
                            '                + intAki1 + intAki2 + intAki3 + intAki4 _
                            '                + intSuikaGenkin + intSuikaCoinLocker + intSonota
                            '.PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_GenkinGokei).Value = intGenkinGokei.ToString("#,##0")
                            ''総売上合計
                            'intSoUriageGokei = intKihonUriageGokei + intGenkinGokei
                            '.PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_SoUriageGokei).Value = intSoUriageGokei.ToString("#,##0")
                            'If intSoUriageGokei < 0 Then
                            '    .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_SoUriageGokei).ForeColor = Color.Red
                            'Else
                            '    .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_SoUriageGokei).ForeColor = Color.Black
                            'End If

                            'intlastRow += 1
                        Next

                        ' --- 最後のデータを書き込む ---
                        '最終行に1行追加
                        .PropvwDayUriageTheatre.Sheets(0).AddUnboundRows(intlastRow, 1)

                        ' 予約番号
                        .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_YoyakuNo).Value = strTempYoyakuNo
                        ' 利用日
                        .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_RiyoDt).Value = strTempYoyakuDt
                        '催事名
                        .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_SaijiNm).Value = strTempSaijiNm
                        '利用者名
                        .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_RiyoNm).Value = strTempRiyoNm
                        '貸出種別
                        .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_KashiKind).Value = strTempKashiKind
                        '利用形状
                        .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_RiyoType).Value = strTempRiyoType
                        '催事分類
                        .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_SaijiBunrui).Value = strTempSaijiBunrui
                        '利用料
                        .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_RiyoKin).Value = lngTempRiyoKin.ToString("#,##0")
                        '付帯設備使用料
                        .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_FutaiKin).Value = lngTempFutaiKin.ToString("#,##0")
                        '付帯設備調整額
                        .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_FutaiChosei).Value = lngTempFutaiChosei.ToString("#,##0")
                        '基本売上合計
                        lngSumUriageGokei = lngTempRiyoKin + lngTempFutaiKin + lngTempFutaiChosei
                        .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_KihonUriageGokei).Value = lngSumUriageGokei.ToString("#,##0")

                        If lngSumUriageGokei < 0 Then
                            ' 合計額がマイナスなら赤文字にする。
                            .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_KihonUriageGokei).ForeColor = Color.Red
                        Else
                            .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_KihonUriageGokei).ForeColor = Color.Black
                        End If

                        'ドリンク現金
                        .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_DrinkGenkin).Value = lngTempDrinkGenkin.ToString("#,##0")
                        'ワンドリンク
                        .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_OneDrink).Value = lngTempOneDrink.ToString("#,##0")
                        'コインロッカー
                        .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_CoinLocker).Value = lngTempCoinLocker.ToString("#,##0")
                        '雑収入
                        .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_ZatsuShunyu).Value = lngTempZatsuShunyu.ToString("#,##0")
                        '空き1
                        .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_Aki1).Value = lngTempAki1.ToString("#,##0")
                        '空き2
                        .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_Aki2).Value = lngTempAki2.ToString("#,##0")
                        '空き3
                        .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_Aki3).Value = lngTempAki3.ToString("#,##0")
                        '空き4
                        .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_Aki4).Value = lngTempAki4.ToString("#,##0")
                        'SUIKA現金
                        .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_SuikaGenkin).Value = lngTempSuicaGenkin.ToString("#,##0")
                        'SUIKAコインロッカー
                        .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_SuikaCoinLocker).Value = lngTempSuicaCoinLocker.ToString("#,##0")
                        'その他
                        .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_Sonota).Value = lngTempSonota.ToString("#,##0")
                        '現金合計
                        lngSumGenkin = lngTempDrinkGenkin + lngTempOneDrink + lngTempCoinLocker + lngTempZatsuShunyu _
                                    + lngTempAki1 + lngTempAki2 + lngTempAki3 + lngTempAki4 _
                                    + lngTempSuicaGenkin + lngTempSuicaCoinLocker + lngTempSonota
                        .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_GenkinGokei).Value = lngSumGenkin.ToString("#,##0")
                        '総売上合計
                        lngSumTotalUriage = lngSumUriageGokei + lngSumGenkin
                        .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_SoUriageGokei).Value = lngSumTotalUriage.ToString("#,##0")
                        If lngSumTotalUriage < 0 Then
                            ' 総売上がマイナスなら赤文字にする。
                            .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_SoUriageGokei).ForeColor = Color.Red
                        Else
                            .PropvwDayUriageTheatre.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Theatre_SoUriageGokei).ForeColor = Color.Black
                        End If
                    End If
                    ' --- 2019/08/23 日別売上一覧機能追加対応 Start E.Okuda@Compass ---

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

                    ' --- 2019/08/23 日別売上一覧機能追加対応 Start E.Okuda@Compass ---
                    If .PropDtDayUriage.Rows.Count > 0 Then
                        ' 書き込み用一時変数にデータを格納する。
                        strTempYoyakuNo = .PropDtDayUriage.Rows(0)("YOYAKU_NO").ToString
                        strTempYoyakuDt = .PropDtDayUriage.Rows(0)("YOYAKU_DT").ToString
                        strTempShutsuenNm = .PropDtDayUriage.Rows(0)("SHUTSUEN_NM").ToString
                        strTempRiyoNm = .PropDtDayUriage.Rows(0)("RIYO_NM").ToString
                        strTempKashiKind = .PropDtDayUriage.Rows(0)("KASHI_KIND").ToString()
                        strTempStudioKbn = .PropDtDayUriage.Rows(0)("STUDIO_KBN").ToString()
                        lngTempRiyoKin = .PropDtDayUriage.Rows(0)("RIYO_KIN")
                        lngTempFutaiKin = .PropDtDayUriage.Rows(0)("FUTAI_SHOKEI")
                        lngTempFutaiChosei = .PropDtDayUriage.Rows(0)("FUTAI_CHOSEI")
                        lngTempDrinkGenkin = .PropDtDayUriage.Rows(0)("DRINK_GENKIN")
                        lngTempOneDrink = .PropDtDayUriage.Rows(0)("ONE_DRINK")
                        lngTempCoinLocker = .PropDtDayUriage.Rows(0)("COIN_LOCKER")
                        lngTempZatsuShunyu = .PropDtDayUriage.Rows(0)("ZATSU_SHUNYU")
                        lngTempAki1 = .PropDtDayUriage.Rows(0)("AKI1")
                        lngTempAki2 = .PropDtDayUriage.Rows(0)("AKI2")
                        lngTempAki3 = .PropDtDayUriage.Rows(0)("AKI3")
                        lngTempAki4 = .PropDtDayUriage.Rows(0)("AKI4")
                        lngTempSuicaGenkin = .PropDtDayUriage.Rows(0)("SUICA_GENKIN")
                        lngTempSuicaCoinLocker = .PropDtDayUriage.Rows(0)("SUICA_COIN_LOCKER")
                        lngTempSonota = .PropDtDayUriage.Rows(0)("SONOTA")
                        ' --- 2019/08/23 日別売上一覧機能追加対応 End E.Okuda@Compass ---

                        '---日別売上一覧表示
                        intlastRow = .PropvwDayUriageStudio.ActiveSheet.RowCount
                        ' --- 2019/08/23 日別売上一覧機能追加対応 Start E.Okuda@Compass ---
                        For intDataRow As Integer = 1 To .PropDtDayUriage.Rows.Count - 1

                            If (.PropDtDayUriage.Rows(intDataRow)("YOYAKU_DT").ToString <> strTempYoyakuDt Or
                                    .PropDtDayUriage.Rows(intDataRow)("SHUTSUEN_NM").ToString <> strTempShutsuenNm) Then
                                ' 最終行に1行追加
                                .PropvwDayUriageStudio.Sheets(0).AddUnboundRows(intlastRow, 1)

                                ' 予約番号
                                .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_YoyakuNo).Value = strTempYoyakuNo
                                ' 利用日
                                .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_RiyoDt).Value = strTempYoyakuDt
                                ' アーティスト名
                                .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_ShutsuenNm).Value = strTempShutsuenNm
                                ' 利用者名
                                .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_RiyoNm).Value = strTempRiyoNm
                                ' 貸出種別
                                .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_KashiKind).Value = strTempKashiKind
                                ' スタジオ
                                .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_Studio).Value = strTempStudioKbn
                                ' 利用料
                                .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_RiyoKin).Value = lngTempRiyoKin.ToString("#,##0")
                                ' 付帯設備使用料
                                .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_FutaiKin).Value = lngTempFutaiKin.ToString("#,##0")
                                ' 付帯設備調整額
                                .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_FutaiChosei).Value = lngTempFutaiChosei.ToString("#,##0")
                                ' 基本売上合計
                                lngSumUriageGokei = lngTempRiyoKin + lngTempFutaiKin + lngTempFutaiChosei
                                .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_KihonUriageGokei).Value = lngSumUriageGokei.ToString("#,##0")
                                If lngSumUriageGokei < 0 Then
                                    .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_KihonUriageGokei).ForeColor = Color.Red
                                Else
                                    .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_KihonUriageGokei).ForeColor = Color.Black
                                End If
                                ' ドリンク現金
                                .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_DrinkGenkin).Value = lngTempDrinkGenkin.ToString("#,##0")
                                ' ワンドリンク
                                .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_OneDrink).Value = lngTempOneDrink.ToString("#,##0")
                                ' コインロッカー
                                .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_CoinLocker).Value = lngTempCoinLocker.ToString("#,##0")
                                ' 雑収入
                                .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_ZatsuShunyu).Value = lngTempZatsuShunyu.ToString("#,##0")
                                ' 空き1
                                .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_Aki1).Value = lngTempAki1.ToString("#,##0")
                                ' 空き2
                                .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_Aki2).Value = lngTempAki2.ToString("#,##0")
                                ' 空き3
                                .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_Aki3).Value = lngTempAki3.ToString("#,##0")
                                ' 空き4
                                .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_Aki4).Value = lngTempAki4.ToString("#,##0")
                                ' SUIKA現金
                                .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_SuikaGenkin).Value = lngTempSuicaGenkin.ToString("#,##0")
                                ' SUIKAコインロッカー
                                .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_SuikaCoinLocker).Value = lngTempSuicaCoinLocker.ToString("#,##0")
                                ' その他
                                .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_Sonota).Value = lngTempSonota.ToString("#,##0")
                                ' 現金合計
                                lngSumGenkin = lngTempDrinkGenkin + lngTempOneDrink + lngTempCoinLocker + lngTempZatsuShunyu _
                                                + lngTempAki1 + lngTempAki2 + lngTempAki3 + lngTempAki4 _
                                                + lngTempSuicaGenkin + lngTempSuicaCoinLocker + lngTempSonota
                                .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_GenkinGokei).Value = lngSumGenkin.ToString("#,##0")
                                ' 総売上合計
                                lngSumTotalUriage = lngSumUriageGokei + lngSumGenkin
                                .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_SoUriageGokei).Value = lngSumTotalUriage.ToString("#,##0")
                                If lngSumTotalUriage < 0 Then
                                    .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_SoUriageGokei).ForeColor = Color.Red
                                Else
                                    .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_SoUriageGokei).ForeColor = Color.Black
                                End If

                                intlastRow += 1

                                ' 書き込み用一時変数にデータを格納する。
                                strTempYoyakuNo = .PropDtDayUriage.Rows(intDataRow)("YOYAKU_NO").ToString
                                strTempYoyakuDt = .PropDtDayUriage.Rows(intDataRow)("YOYAKU_DT").ToString
                                strTempShutsuenNm = .PropDtDayUriage.Rows(intDataRow)("SHUTSUEN_NM").ToString
                                strTempRiyoNm = .PropDtDayUriage.Rows(intDataRow)("RIYO_NM").ToString
                                strTempKashiKind = .PropDtDayUriage.Rows(intDataRow)("KASHI_KIND").ToString()
                                strTempStudioKbn = .PropDtDayUriage.Rows(intDataRow)("STUDIO_KBN").ToString()
                                lngTempRiyoKin = .PropDtDayUriage.Rows(intDataRow)("RIYO_KIN")
                                lngTempFutaiKin = .PropDtDayUriage.Rows(intDataRow)("FUTAI_SHOKEI")
                                lngTempFutaiChosei = .PropDtDayUriage.Rows(intDataRow)("FUTAI_CHOSEI")
                                lngTempDrinkGenkin = .PropDtDayUriage.Rows(intDataRow)("DRINK_GENKIN")
                                lngTempOneDrink = .PropDtDayUriage.Rows(intDataRow)("ONE_DRINK")
                                lngTempCoinLocker = .PropDtDayUriage.Rows(intDataRow)("COIN_LOCKER")
                                lngTempZatsuShunyu = .PropDtDayUriage.Rows(intDataRow)("ZATSU_SHUNYU")
                                lngTempAki1 = .PropDtDayUriage.Rows(intDataRow)("AKI1")
                                lngTempAki2 = .PropDtDayUriage.Rows(intDataRow)("AKI2")
                                lngTempAki3 = .PropDtDayUriage.Rows(intDataRow)("AKI3")
                                lngTempAki4 = .PropDtDayUriage.Rows(intDataRow)("AKI4")
                                lngTempSuicaGenkin = .PropDtDayUriage.Rows(intDataRow)("SUICA_GENKIN")
                                lngTempSuicaCoinLocker = .PropDtDayUriage.Rows(intDataRow)("SUICA_COIN_LOCKER")
                                lngTempSonota = .PropDtDayUriage.Rows(intDataRow)("SONOTA")
                            Else
                                ' 予約番号
                                If Not strTempYoyakuNo = .PropDtDayUriage.Rows(intDataRow)("YOYAKU_NO").ToString Then
                                    strTempYoyakuNo = strTempYoyakuNo & STR_DELIMITER & .PropDtDayUriage.Rows(intDataRow)("YOYAKU_NO").ToString
                                End If
                                ' 利用日
                                If Not strTempYoyakuDt = .PropDtDayUriage.Rows(intDataRow)("YOYAKU_DT").ToString Then
                                    strTempYoyakuDt = strTempYoyakuDt & STR_DELIMITER & .PropDtDayUriage.Rows(intDataRow)("YOYAKU_DT").ToString
                                End If
                                ' アーティスト名
                                If Not strTempShutsuenNm = .PropDtDayUriage.Rows(intDataRow)("SHUTSUEN_NM").ToString Then
                                    strTempShutsuenNm = strTempShutsuenNm & STR_DELIMITER & .PropDtDayUriage.Rows(intDataRow)("SHUTSUEN_NM").ToString
                                End If
                                ' 利用者名
                                If Not strTempRiyoNm = .PropDtDayUriage.Rows(intDataRow)("RIYO_NM").ToString Then
                                    strTempRiyoNm = strTempRiyoNm & STR_DELIMITER & .PropDtDayUriage.Rows(intDataRow)("RIYO_NM").ToString
                                End If
                                ' 貸出種別
                                If Not strTempKashiKind = .PropDtDayUriage.Rows(intDataRow)("KASHI_KIND").ToString() Then
                                    strTempKashiKind = strTempKashiKind & STR_DELIMITER & .PropDtDayUriage.Rows(intDataRow)("KASHI_KIND").ToString()
                                End If
                                ' スタジオ
                                If Not strTempStudioKbn = .PropDtDayUriage.Rows(intDataRow)("STUDIO_KBN").ToString() Then
                                    strTempStudioKbn = strTempStudioKbn & STR_DELIMITER & .PropDtDayUriage.Rows(intDataRow)("STUDIO_KBN").ToString()
                                End If
                                ' 利用料
                                lngTempRiyoKin += .PropDtDayUriage.Rows(intDataRow)("RIYO_KIN")
                                ' 付帯設備使用料
                                lngTempFutaiKin += .PropDtDayUriage.Rows(intDataRow)("FUTAI_SHOKEI")
                                ' 付帯設備調整額
                                lngTempFutaiChosei += .PropDtDayUriage.Rows(intDataRow)("FUTAI_CHOSEI")
                                ' ドリンク現金
                                lngTempDrinkGenkin += .PropDtDayUriage.Rows(intDataRow)("DRINK_GENKIN")
                                ' ワンドリンク
                                lngTempOneDrink += .PropDtDayUriage.Rows(intDataRow)("ONE_DRINK")
                                ' コインロッカー
                                lngTempCoinLocker += .PropDtDayUriage.Rows(intDataRow)("COIN_LOCKER")
                                ' 雑収入
                                lngTempZatsuShunyu += .PropDtDayUriage.Rows(intDataRow)("ZATSU_SHUNYU")
                                ' 空き1
                                lngTempAki1 += .PropDtDayUriage.Rows(intDataRow)("AKI1")
                                ' 空き2
                                lngTempAki2 += .PropDtDayUriage.Rows(intDataRow)("AKI2")
                                ' 空き3
                                lngTempAki3 += .PropDtDayUriage.Rows(intDataRow)("AKI3")
                                ' 空き4
                                lngTempAki4 += .PropDtDayUriage.Rows(intDataRow)("AKI4")
                                ' SUIKA現金
                                lngTempSuicaGenkin += .PropDtDayUriage.Rows(intDataRow)("SUICA_GENKIN")
                                ' SUIKAコインロッカー
                                lngTempSuicaCoinLocker += .PropDtDayUriage.Rows(intDataRow)("SUICA_COIN_LOCKER")
                                ' その他
                                lngTempSonota += .PropDtDayUriage.Rows(intDataRow)("SONOTA")

                            End If


                            ' For intDataRow As Integer = 0 To .PropDtDayUriage.Rows.Count - 1
                            ''最終行に1行追加
                            '.PropvwDayUriageStudio.Sheets(0).AddUnboundRows(intlastRow, 1)

                            ''予約番号
                            '.PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_YoyakuNo).Value = .PropDtDayUriage.Rows(intDataRow)("YOYAKU_NO").ToString
                            ''利用日
                            '.PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_RiyoDt).Value = .PropDtDayUriage.Rows(intDataRow)("YOYAKU_DT").ToString
                            ''アーティスト名
                            '.PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_ShutsuenNm).Value = .PropDtDayUriage.Rows(intDataRow)("SHUTSUEN_NM").ToString
                            ''利用者名
                            '.PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_RiyoNm).Value = .PropDtDayUriage.Rows(intDataRow)("RIYO_NM").ToString
                            ''貸出種別
                            '.PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_KashiKind).Value = .PropDtDayUriage.Rows(intDataRow)("KASHI_KIND").ToString
                            ''スタジオ
                            '.PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_Studio).Value = .PropDtDayUriage.Rows(intDataRow)("STUDIO_KBN").ToString
                            ''利用料
                            'intRiyoKin = .PropDtDayUriage.Rows(intDataRow)("RIYO_KIN")
                            '.PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_RiyoKin).Value = intRiyoKin.ToString("#,##0")
                            ''付帯設備使用料
                            'intFutaiKin = .PropDtDayUriage.Rows(intDataRow)("FUTAI_KIN")
                            '.PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_FutaiKin).Value = intFutaiKin.ToString("#,##0")
                            ''付帯設備調整額
                            'intFutaiChosei = .PropDtDayUriage.Rows(intDataRow)("FUTAI_CHOSEI")
                            '.PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_FutaiChosei).Value = intFutaiChosei.ToString("#,##0")
                            ''基本売上合計
                            'intKihonUriageGokei = intRiyoKin + intFutaiKin + intFutaiChosei
                            '.PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_KihonUriageGokei).Value = intKihonUriageGokei.ToString("#,##0")
                            'If intKihonUriageGokei < 0 Then
                            '    .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_KihonUriageGokei).ForeColor = Color.Red
                            'Else
                            '    .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_KihonUriageGokei).ForeColor = Color.Black
                            'End If
                            ''ドリンク現金
                            'intDrinkGenkin = .PropDtDayUriage.Rows(intDataRow)("DRINK_GENKIN")
                            '.PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_DrinkGenkin).Value = intDrinkGenkin.ToString("#,##0")
                            ''ワンドリンク
                            'intOneDrink = .PropDtDayUriage.Rows(intDataRow)("ONE_DRINK")
                            '.PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_OneDrink).Value = intOneDrink.ToString("#,##0")
                            ''コインロッカー
                            'intCoinLocker = .PropDtDayUriage.Rows(intDataRow)("COIN_LOCKER")
                            '.PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_CoinLocker).Value = intCoinLocker.ToString("#,##0")
                            ''雑収入
                            'intZatsuShunyu = .PropDtDayUriage.Rows(intDataRow)("ZATSU_SHUNYU")
                            '.PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_ZatsuShunyu).Value = intZatsuShunyu.ToString("#,##0")
                            ''空き1
                            'intAki1 = .PropDtDayUriage.Rows(intDataRow)("AKI1")
                            '.PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_Aki1).Value = intAki1.ToString("#,##0")
                            ''空き2
                            'intAki2 = .PropDtDayUriage.Rows(intDataRow)("AKI2")
                            '.PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_Aki2).Value = intAki2.ToString("#,##0")
                            ''空き3
                            'intAki3 = .PropDtDayUriage.Rows(intDataRow)("AKI3")
                            '.PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_Aki3).Value = intAki3.ToString("#,##0")
                            ''空き4
                            'intAki4 = .PropDtDayUriage.Rows(intDataRow)("AKI4")
                            '.PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_Aki4).Value = intAki4.ToString("#,##0")
                            ''SUIKA現金
                            'intSuikaGenkin = .PropDtDayUriage.Rows(intDataRow)("SUICA_GENKIN")
                            '.PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_SuikaGenkin).Value = intSuikaGenkin.ToString("#,##0")
                            ''SUIKAコインロッカー
                            'intSuikaCoinLocker = .PropDtDayUriage.Rows(intDataRow)("SUICA_COIN_LOCKER")
                            '.PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_SuikaCoinLocker).Value = intSuikaCoinLocker.ToString("#,##0")
                            ''その他
                            'intSonota = .PropDtDayUriage.Rows(intDataRow)("SONOTA")
                            '.PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_Sonota).Value = intSonota.ToString("#,##0")
                            ''現金合計
                            'intGenkinGokei = intDrinkGenkin + intOneDrink + intCoinLocker + intZatsuShunyu _
                            '                + intAki1 + intAki2 + intAki3 + intAki4 _
                            '                + intSuikaGenkin + intSuikaCoinLocker + intSonota
                            '.PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_GenkinGokei).Value = intGenkinGokei.ToString("#,##0")
                            ''総売上合計
                            'intSoUriageGokei = intKihonUriageGokei + intGenkinGokei
                            '.PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_SoUriageGokei).Value = intSoUriageGokei.ToString("#,##0")
                            'If intSoUriageGokei < 0 Then
                            '    .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_SoUriageGokei).ForeColor = Color.Red
                            'Else
                            '    .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_SoUriageGokei).ForeColor = Color.Black
                            'End If

                            'intlastRow += 1
                            'Next

                        Next

                        ' --- 最後のデータを書き込む ---
                        '最終行に1行追加
                        .PropvwDayUriageStudio.Sheets(0).AddUnboundRows(intlastRow, 1)

                        ' 予約番号
                        .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_YoyakuNo).Value = strTempYoyakuNo
                        ' 利用日
                        .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_RiyoDt).Value = strTempYoyakuDt
                        ' アーティスト名
                        .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_ShutsuenNm).Value = strTempShutsuenNm
                        ' 利用者名
                        .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_RiyoNm).Value = strTempRiyoNm
                        ' 貸出種別
                        .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_KashiKind).Value = strTempKashiKind
                        ' スタジオ
                        .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_Studio).Value = strTempStudioKbn
                        ' 利用料
                        .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_RiyoKin).Value = lngTempRiyoKin.ToString("#,##0")
                        ' 付帯設備使用料
                        .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_FutaiKin).Value = lngTempFutaiKin.ToString("#,##0")
                        ' 付帯設備調整額
                        .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_FutaiChosei).Value = lngTempFutaiChosei.ToString("#,##0")
                        ' 基本売上合計
                        lngSumUriageGokei = lngTempRiyoKin + lngTempFutaiKin + lngTempFutaiChosei
                        .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_KihonUriageGokei).Value = lngSumUriageGokei.ToString("#,##0")
                        If lngSumUriageGokei < 0 Then
                            .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_KihonUriageGokei).ForeColor = Color.Red
                        Else
                            .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_KihonUriageGokei).ForeColor = Color.Black
                        End If
                        ' ドリンク現金
                        .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_DrinkGenkin).Value = lngTempDrinkGenkin.ToString("#,##0")
                        ' ワンドリンク
                        .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_OneDrink).Value = lngTempOneDrink.ToString("#,##0")
                        ' コインロッカー
                        .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_CoinLocker).Value = lngTempCoinLocker.ToString("#,##0")
                        ' 雑収入
                        .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_ZatsuShunyu).Value = lngTempZatsuShunyu.ToString("#,##0")
                        ' 空き1
                        .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_Aki1).Value = lngTempAki1.ToString("#,##0")
                        ' 空き2
                        .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_Aki2).Value = lngTempAki2.ToString("#,##0")
                        ' 空き3
                        .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_Aki3).Value = lngTempAki3.ToString("#,##0")
                        ' 空き4
                        .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_Aki4).Value = lngTempAki4.ToString("#,##0")
                        ' SUIKA現金
                        .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_SuikaGenkin).Value = lngTempSuicaGenkin.ToString("#,##0")
                        ' SUIKAコインロッカー
                        .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_SuikaCoinLocker).Value = lngTempSuicaCoinLocker.ToString("#,##0")
                        ' その他
                        .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_Sonota).Value = lngTempSonota.ToString("#,##0")
                        ' 現金合計
                        lngSumGenkin = lngTempDrinkGenkin + lngTempOneDrink + lngTempCoinLocker + lngTempZatsuShunyu _
                                                + lngTempAki1 + lngTempAki2 + lngTempAki3 + lngTempAki4 _
                                                + lngTempSuicaGenkin + lngTempSuicaCoinLocker + lngTempSonota
                        .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_GenkinGokei).Value = lngSumGenkin.ToString("#,##0")
                        ' 総売上合計
                        lngSumTotalUriage = lngSumUriageGokei + lngSumGenkin
                        .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_SoUriageGokei).Value = lngSumTotalUriage.ToString("#,##0")
                        If lngSumTotalUriage < 0 Then
                            .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_SoUriageGokei).ForeColor = Color.Red
                        Else
                            .PropvwDayUriageStudio.ActiveSheet.Cells(intlastRow, SpreadDayUriageIndex_Studio_SoUriageGokei).ForeColor = Color.Black
                        End If
                    End If
                    ' --- 2019/08/23 日別売上一覧機能追加対応 End E.Okuda@Compass ---

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

    ' --- 2019/08/23 日別売上一覧機能追加対応 Start E.Okuda@Compass ---
    ''' <summary>
    ''' 集計CSV出力メイン処理
    ''' </summary>
    ''' <param name="dataEXTZ0103"></param>
    ''' <returns></returns>

    Public Function OutputCsvMain(ByVal dataEXTZ0103 As DataEXTZ0103) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Try
            'CSV出力するか確認メッセージ
            If MsgBox(Z0103_C0001, MsgBoxStyle.OkCancel, TITLE_INFO) = vbOK Then
                If OutputCsvData(dataEXTZ0103) = False Then
                    Return False
                End If
            End If

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

    ''' <summary>
    ''' 日別売上一覧CSV出力
    ''' </summary>
    ''' <param name="dataEXTZ0103">データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了　false 異常終了</returns>
    ''' <remarks>スプレッドのデータをCSV出力する。
    ''' <para>作成情報：2018/08/27 E.Okuda@Compass
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Private Function OutputCsvData(ByVal dataEXTZ0103 As DataEXTZ0103) As Boolean
        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Const STR_DOUBLE_QUOTATION = """"                                               ' ダブルクォーテーション
        Const STR_COMMA As String = ","                                                 ' カンマ
        Const COL_IDX_RIYOSYA As Integer = 3
        Dim strFileName As String                                                       ' 出力ファイル名
        Dim strFilePath As String = ""                                                  ' 出力ファイルパス
        Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding("Shift_JIS") ' 出力ファイルの文字コード
        Dim intRowIdx As Integer                                                        ' スプレッドシート行インデックス
        Dim intColIdx As Integer                                                        ' スプレッドシートカラムインデックス


        Try
            ' デフォルトファイル名設定
            strFileName = "日別売上_" & DateTime.Now.ToString("yyyyMMddHHmmss")

            ' ファイル保存ダイアログ起動
            Using sfd As New SaveFileDialog()
                sfd.Filter = "日別売上データ(*.csv)|*.csv"
                sfd.FileName = strFileName

                If sfd.ShowDialog() = DialogResult.OK Then
                    'OKが押下された場合
                    'ファイル名を設定
                    strFilePath = sfd.FileName

                    '書き込むファイルを開く
                    Dim swCsvFile As New System.IO.StreamWriter(strFilePath, False, enc)

                    If dataEXTZ0103.PropvwDayUriageTheatre.Visible Then
                        ' シアタースプレッド表示時

                        'ヘッダを書き込む
                        swCsvFile.Write(CSV_HEADER_THEATER)
                        '改行する
                        swCsvFile.Write(vbCrLf)

                        With dataEXTZ0103.PropvwDayUriageTheatre.Sheets(0)
                            For intRowIdx = 0 To .RowCount - 1
                                For intColIdx = 0 To .ColumnCount - 1
                                    If intColIdx <> COL_IDX_RIYOSYA Then
                                        swCsvFile.Write(STR_DOUBLE_QUOTATION & dataEXTZ0103.PropvwDayUriageTheatre.Sheets(0).Cells(intRowIdx, intColIdx).Value & STR_DOUBLE_QUOTATION)

                                        ' 最終列は改行コード、それ以外はカンマを付ける
                                        If intColIdx < .ColumnCount - 1 Then
                                            swCsvFile.Write(STR_COMMA)
                                        Else
                                            swCsvFile.Write(vbCrLf)
                                        End If

                                    End If

                                Next
                            Next
                        End With

                    Else
                        ' スタジオスプレッド表示時
                        'ヘッダを書き込む
                        swCsvFile.Write(CSV_HEADER_STUDIO)
                        '改行する
                        swCsvFile.Write(vbCrLf)

                        With dataEXTZ0103.PropvwDayUriageStudio.Sheets(0)
                            For intRowIdx = 0 To .RowCount - 1
                                For intColIdx = 0 To .ColumnCount - 1
                                    If intColIdx <> COL_IDX_RIYOSYA Then
                                        swCsvFile.Write(STR_DOUBLE_QUOTATION & dataEXTZ0103.PropvwDayUriageStudio.Sheets(0).Cells(intRowIdx, intColIdx).Value & STR_DOUBLE_QUOTATION)

                                        ' 最終列は改行コード、それ以外はカンマを付ける
                                        If intColIdx < .ColumnCount - 1 Then
                                            swCsvFile.Write(STR_COMMA)
                                        Else
                                            swCsvFile.Write(vbCrLf)
                                        End If

                                    End If

                                Next
                            Next
                        End With

                    End If

                    'CSVファイルを閉じる
                    swCsvFile.Close()

                    '出力処理正常終了時、メッセージを出力
                    MsgBox(String.Format(Z0103_I0001, "CSVファイル出力"))

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
    ' --- 2019/08/23 日別売上一覧機能追加対応 End E.Okuda@Compass ---


#End Region

End Class
