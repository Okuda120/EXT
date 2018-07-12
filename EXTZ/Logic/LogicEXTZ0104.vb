Imports Common
Imports CommonEXT
Imports Npgsql

''' <summary>
''' LogicEXTZ0104
''' </summary>
''' <remarks>使用状況一覧画面で発生する情報の取得処理等を行う
''' <para>作成情報：2015/09/24 h.endo
''' <p>改訂情報:</p>
''' </para></remarks>
Public Class LogicEXTZ0104

    '変数宣言
    Private sqlEXTZ0104 As New SqlEXTZ0104              'sqlクラス

#Region "スプレッドの作成"

    ''' <summary>
    ''' スプレッドの作成
    ''' </summary>
    ''' <param name="dataEXTZ0104">データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了　false 異常終了</returns>
    ''' <remarks>スプレッドを作成する。
    ''' <para>作成情報：2015/09/24 h.endo
    ''' <p>改訂情報:2015/10/20 m.hayabuchi</p>
    ''' </para></remarks>
    Public Function MakeSpread(ByRef dataEXTZ0104 As DataEXTZ0104) As Boolean

        '変数宣言
        Dim dtRiyoJokyo As DataTable = Nothing          '月毎のデータ
        Dim intTotal As Integer = 0                     '使用日数（使用形態）
        Dim intStotal As Integer = 0                    '使用日数（催事分類）
        Dim intSetTsuki As Integer = 0                  'ヘッダに設定する月
        Dim intGokKano As Integer = 0                   '可能日数（合計）
        Dim intGokKyukan As Integer = 0                 '休館日数（合計）
        Dim intGokMainte As Integer = 0                 'メンテ日数（合計）
        Dim intGokTotal As Integer = 0                  '使用日数（合計）
        Dim intGokChakuseki As Integer = 0              '着席使用日数（合計）
        Dim intGokStanding As Integer = 0               'スタンディング使用日数（合計）
        Dim intGokMix As Integer = 0                    'MIX使用日数（合計）
        Dim intGokSaiji As Integer = 0                  '催事使用日数（合計）
        Dim intGokMusic As Integer = 0                  '音楽使用日数（合計）
        Dim intGokEngeki As Integer = 0                 '演劇使用日数（合計）
        Dim intGokEngei As Integer = 0                  '演芸使用日数（合計）
        Dim intGokBusiness As Integer = 0               'ビジネス使用日数（合計）
        Dim intGokMovie As Integer = 0                  '試写会・映画使用日数（合計）
        Dim intGokSonota As Integer = 0                 'その他使用日数（合計）
        Dim intGok201st As Integer = 0                  '201st使用日数（合計）
        Dim intGok202st As Integer = 0                  '202st使用日数（合計）
        Dim intGokhouselock As Integer = 0              'houselock使用日数（合計）

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try

            With dataEXTZ0104

                'データ取得
                If GetSpreadData(dataEXTZ0104) = False Then
                    Return False
                End If

                '一覧表示
                intGokKano = 0
                intGokKyukan = 0
                intGokMainte = 0
                intGokTotal = 0
                intGokChakuseki = 0
                intGokStanding = 0
                intGokMix = 0
                intGokSaiji = 0
                intGokMusic = 0
                intGokEngeki = 0
                intGokEngei = 0
                intGokBusiness = 0
                intGokMovie = 0
                intGokSonota = 0
                intGok201st = 0
                intGok202st = 0
                intGokhouselock = 0

                For intTsukiCnt As Integer = 1 To 12
                    Select Case intTsukiCnt
                        Case 1
                            dtRiyoJokyo = .PropDtRiyoJokyo1
                        Case 2
                            dtRiyoJokyo = .PropDtRiyoJokyo2
                        Case 3
                            dtRiyoJokyo = .PropDtRiyoJokyo3
                        Case 4
                            dtRiyoJokyo = .PropDtRiyoJokyo4
                        Case 5
                            dtRiyoJokyo = .PropDtRiyoJokyo5
                        Case 6
                            dtRiyoJokyo = .PropDtRiyoJokyo6
                        Case 7
                            dtRiyoJokyo = .PropDtRiyoJokyo7
                        Case 8
                            dtRiyoJokyo = .PropDtRiyoJokyo8
                        Case 9
                            dtRiyoJokyo = .PropDtRiyoJokyo9
                        Case 10
                            dtRiyoJokyo = .PropDtRiyoJokyo10
                        Case 11
                            dtRiyoJokyo = .PropDtRiyoJokyo11
                        Case 12
                            dtRiyoJokyo = .PropDtRiyoJokyo12
                    End Select

                    intTotal = 0
                    intStotal = 0


                    If .PropStrShisetsuKbn = SHISETU_KBN_THEATER Then
                        'シアターの場合

                        '---利用状況一覧表示
                        For intDataRow As Integer = 0 To dtRiyoJokyo.Rows.Count - 1

                            intSetTsuki = Integer.Parse(.PropStrKikanTsuki) + (intTsukiCnt - 1)
                            If intSetTsuki > 12 Then
                                intSetTsuki -= 12
                            End If
                            .PropvwRiyoJokyoTheatre.Sheets(0).ColumnHeader.Cells(0, intTsukiCnt - 1).Value = intSetTsuki.ToString & "月"
                            .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_KANO_NISSU, intTsukiCnt - 1).Value = dtRiyoJokyo.Rows(intDataRow)("RIYO_KANO_NISSU").ToString
                            intGokKano += Integer.Parse(dtRiyoJokyo.Rows(intDataRow)("RIYO_KANO_NISSU"))
                            .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_KYUKAN_NISSU, intTsukiCnt - 1).Value = dtRiyoJokyo.Rows(intDataRow)("KYUKAN_NISSU").ToString
                            intGokKyukan += Integer.Parse(dtRiyoJokyo.Rows(intDataRow)("KYUKAN_NISSU"))
                            .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_MAINTE_NISSU, intTsukiCnt - 1).Value = dtRiyoJokyo.Rows(intDataRow)("MAINTE_NISSU").ToString
                            intGokMainte += Integer.Parse(dtRiyoJokyo.Rows(intDataRow)("MAINTE_NISSU"))

                            .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_NISSU_CHAKUSEKI, intTsukiCnt - 1).Value = dtRiyoJokyo.Rows(intDataRow)("RIYO_NISSU_CHAKUSEKI").ToString
                            intTotal += Integer.Parse(dtRiyoJokyo.Rows(intDataRow)("RIYO_NISSU_CHAKUSEKI"))
                            intGokChakuseki += Integer.Parse(dtRiyoJokyo.Rows(intDataRow)("RIYO_NISSU_CHAKUSEKI"))
                            .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_NISSU_STANDING, intTsukiCnt - 1).Value = dtRiyoJokyo.Rows(intDataRow)("RIYO_NISSU_STANDING").ToString
                            intTotal += Integer.Parse(dtRiyoJokyo.Rows(intDataRow)("RIYO_NISSU_STANDING"))
                            intGokStanding += Integer.Parse(dtRiyoJokyo.Rows(intDataRow)("RIYO_NISSU_STANDING"))
                            .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_NISSU_MIX, intTsukiCnt - 1).Value = dtRiyoJokyo.Rows(intDataRow)("RIYO_NISSU_MIX").ToString
                            intTotal += Integer.Parse(dtRiyoJokyo.Rows(intDataRow)("RIYO_NISSU_MIX"))
                            intGokMix += Integer.Parse(dtRiyoJokyo.Rows(intDataRow)("RIYO_NISSU_MIX"))
                            .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_NISSU_SAIJI, intTsukiCnt - 1).Value = dtRiyoJokyo.Rows(intDataRow)("RIYO_NISSU_SAIJI").ToString
                            intTotal += Integer.Parse(dtRiyoJokyo.Rows(intDataRow)("RIYO_NISSU_SAIJI"))
                            intGokSaiji += Integer.Parse(dtRiyoJokyo.Rows(intDataRow)("RIYO_NISSU_SAIJI"))

                            .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_NISSU_MUSIC, intTsukiCnt - 1).Value = dtRiyoJokyo.Rows(intDataRow)("RIYO_NISSU_MUSIC").ToString
                            intStotal += Integer.Parse(dtRiyoJokyo.Rows(intDataRow)("RIYO_NISSU_MUSIC"))
                            intGokMusic += Integer.Parse(dtRiyoJokyo.Rows(intDataRow)("RIYO_NISSU_MUSIC"))
                            .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_NISSU_ENGEKI, intTsukiCnt - 1).Value = dtRiyoJokyo.Rows(intDataRow)("RIYO_NISSU_ENGEKI").ToString
                            intStotal += Integer.Parse(dtRiyoJokyo.Rows(intDataRow)("RIYO_NISSU_ENGEKI"))
                            ' 2016.01.07 UPD START↓ h.hagiwara 
                            'intGokEngeki = Integer.Parse(dtRiyoJokyo.Rows(intDataRow)("RIYO_NISSU_ENGEKI"))
                            intGokEngeki += Integer.Parse(dtRiyoJokyo.Rows(intDataRow)("RIYO_NISSU_ENGEKI"))
                            ' 2016.01.07 UPD END↑ h.hagiwara 
                            .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_NISSU_ENGEI, intTsukiCnt - 1).Value = dtRiyoJokyo.Rows(intDataRow)("RIYO_NISSU_ENGEI").ToString
                            intStotal += Integer.Parse(dtRiyoJokyo.Rows(intDataRow)("RIYO_NISSU_ENGEI"))
                            intGokEngei += Integer.Parse(dtRiyoJokyo.Rows(intDataRow)("RIYO_NISSU_ENGEI"))
                            .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_NISSU_BUSINESS, intTsukiCnt - 1).Value = dtRiyoJokyo.Rows(intDataRow)("RIYO_NISSU_BUSINESS").ToString
                            intStotal += Integer.Parse(dtRiyoJokyo.Rows(intDataRow)("RIYO_NISSU_BUSINESS"))
                            intGokBusiness += Integer.Parse(dtRiyoJokyo.Rows(intDataRow)("RIYO_NISSU_BUSINESS"))
                            .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_NISSU_MOVIE, intTsukiCnt - 1).Value = dtRiyoJokyo.Rows(intDataRow)("RIYO_NISSU_MOVIE").ToString
                            intStotal += Integer.Parse(dtRiyoJokyo.Rows(intDataRow)("RIYO_NISSU_MOVIE"))
                            intGokMovie += Integer.Parse(dtRiyoJokyo.Rows(intDataRow)("RIYO_NISSU_MOVIE"))
                            .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_NISSU_SONOTA, intTsukiCnt - 1).Value = dtRiyoJokyo.Rows(intDataRow)("RIYO_NISSU_SONOTA").ToString
                            intStotal += Integer.Parse(dtRiyoJokyo.Rows(intDataRow)("RIYO_NISSU_SONOTA"))
                            intGokSonota += Integer.Parse(dtRiyoJokyo.Rows(intDataRow)("RIYO_NISSU_SONOTA"))

                            .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_NISSU_TOTAL, intTsukiCnt - 1).Value = intTotal.ToString
                            intGokTotal += intTotal

                            If Integer.Parse(dtRiyoJokyo.Rows(intDataRow)("RIYO_KANO_NISSU")) <> 0 Then
                                .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_TOTAL, intTsukiCnt - 1).Value = _
                                Math.Round(intTotal * 100 / Integer.Parse(dtRiyoJokyo.Rows(intDataRow)("RIYO_KANO_NISSU")), 0)
                            Else
                                .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_TOTAL, intTsukiCnt - 1).Value = "0"
                            End If
                            If intTotal <> 0 Then
                                .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_CHAKUSEKI, intTsukiCnt - 1).Value = _
                                Math.Round(Integer.Parse(dtRiyoJokyo.Rows(intDataRow)("RIYO_NISSU_CHAKUSEKI")) * 100 / intTotal, 0)
                                .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_STANDING, intTsukiCnt - 1).Value = _
                                Math.Round(Integer.Parse(dtRiyoJokyo.Rows(intDataRow)("RIYO_NISSU_STANDING")) * 100 / intTotal, 0)
                                .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_MIX, intTsukiCnt - 1).Value = _
                                Math.Round(Integer.Parse(dtRiyoJokyo.Rows(intDataRow)("RIYO_NISSU_MIX")) * 100 / intTotal, 0)
                                .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_SAIJI, intTsukiCnt - 1).Value = _
                                Math.Round(Integer.Parse(dtRiyoJokyo.Rows(intDataRow)("RIYO_NISSU_SAIJI")) * 100 / intTotal, 0)
                                .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_MUSIC, intTsukiCnt - 1).Value = _
                                Math.Round(Integer.Parse(dtRiyoJokyo.Rows(intDataRow)("RIYO_NISSU_MUSIC")) * 100 / intTotal, 0)
                                .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_ENGEKI, intTsukiCnt - 1).Value = _
                                Math.Round(Integer.Parse(dtRiyoJokyo.Rows(intDataRow)("RIYO_NISSU_ENGEKI")) * 100 / intTotal, 0)
                                .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_ENGEI, intTsukiCnt - 1).Value = _
                                Math.Round(Integer.Parse(dtRiyoJokyo.Rows(intDataRow)("RIYO_NISSU_ENGEI")) * 100 / intTotal, 0)
                                .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_BUSINESS, intTsukiCnt - 1).Value = _
                                Math.Round(Integer.Parse(dtRiyoJokyo.Rows(intDataRow)("RIYO_NISSU_BUSINESS")) * 100 / intTotal, 0)
                                .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_MOVIE, intTsukiCnt - 1).Value = _
                                Math.Round(Integer.Parse(dtRiyoJokyo.Rows(intDataRow)("RIYO_NISSU_MOVIE")) * 100 / intTotal, 0)
                                .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_SONOTA, intTsukiCnt - 1).Value = _
                                Math.Round(Integer.Parse(dtRiyoJokyo.Rows(intDataRow)("RIYO_NISSU_SONOTA")) * 100 / intTotal, 0)
                            Else
                                .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_CHAKUSEKI, intTsukiCnt - 1).Value = "0"
                                .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_STANDING, intTsukiCnt - 1).Value = "0"
                                .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_MIX, intTsukiCnt - 1).Value = "0"
                                .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_SAIJI, intTsukiCnt - 1).Value = "0"
                                .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_MUSIC, intTsukiCnt - 1).Value = "0"
                                .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_ENGEKI, intTsukiCnt - 1).Value = "0"
                                .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_ENGEI, intTsukiCnt - 1).Value = "0"
                                .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_BUSINESS, intTsukiCnt - 1).Value = "0"
                                .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_MOVIE, intTsukiCnt - 1).Value = "0"
                                .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_SONOTA, intTsukiCnt - 1).Value = "0"
                            End If
                        Next

                        If intTsukiCnt = 12 Then
                            '合計列表示
                            .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_KANO_NISSU, 12).Value = intGokKano.ToString
                            .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_KYUKAN_NISSU, 12).Value = intGokKyukan.ToString
                            .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_MAINTE_NISSU, 12).Value = intGokMainte.ToString
                            .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_NISSU_TOTAL, 12).Value = intGokTotal.ToString

                            If intGokKano <> 0 Then
                                .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_TOTAL, 12).Value = Math.Round(intGokTotal * 100 / intGokKano, 0).ToString
                            Else
                                .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_TOTAL, 12).Value = "0"
                            End If

                            .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_NISSU_CHAKUSEKI, 12).Value = intGokChakuseki.ToString
                            .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_NISSU_STANDING, 12).Value = intGokStanding.ToString
                            .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_NISSU_MIX, 12).Value = intGokMix.ToString
                            .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_NISSU_SAIJI, 12).Value = intGokSaiji.ToString
                            .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_NISSU_MUSIC, 12).Value = intGokMusic.ToString
                            .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_NISSU_ENGEKI, 12).Value = intGokEngeki.ToString
                            .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_NISSU_ENGEI, 12).Value = intGokEngei.ToString
                            .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_NISSU_BUSINESS, 12).Value = intGokBusiness.ToString
                            .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_NISSU_MOVIE, 12).Value = intGokMovie.ToString
                            .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_NISSU_SONOTA, 12).Value = intGokSonota.ToString

                            If intGokTotal <> 0 Then
                                .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_CHAKUSEKI, 12).Value = Math.Round(intGokChakuseki * 100 / intGokTotal, 0).ToString
                                .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_STANDING, 12).Value = Math.Round(intGokStanding * 100 / intGokTotal, 0).ToString
                                .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_MIX, 12).Value = Math.Round(intGokMix * 100 / intGokTotal, 0).ToString
                                .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_SAIJI, 12).Value = Math.Round(intGokSaiji * 100 / intGokTotal, 0).ToString
                                .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_MUSIC, 12).Value = Math.Round(intGokMusic * 100 / intGokTotal, 0).ToString
                                .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_ENGEKI, 12).Value = Math.Round(intGokEngeki * 100 / intGokTotal, 0).ToString
                                .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_ENGEI, 12).Value = Math.Round(intGokEngei * 100 / intGokTotal, 0).ToString
                                .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_BUSINESS, 12).Value = Math.Round(intGokBusiness * 100 / intGokTotal, 0).ToString
                                .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_MOVIE, 12).Value = Math.Round(intGokMovie * 100 / intGokTotal, 0).ToString
                                .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_SONOTA, 12).Value = Math.Round(intGokSonota * 100 / intGokTotal, 0).ToString
                            Else
                                .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_CHAKUSEKI, 12).Value = "0"
                                .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_STANDING, 12).Value = "0"
                                .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_MIX, 12).Value = "0"
                                .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_SAIJI, 12).Value = "0"
                                .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_MUSIC, 12).Value = "0"
                                .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_ENGEKI, 12).Value = "0"
                                .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_ENGEI, 12).Value = "0"
                                .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_BUSINESS, 12).Value = "0"
                                .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_MOVIE, 12).Value = "0"
                                .PropvwRiyoJokyoTheatre.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_SONOTA, 12).Value = "0"
                            End If

                        End If

                    Else
                        'スタジオの場合

                        For intDataRow As Integer = 0 To dtRiyoJokyo.Rows.Count - 1

                            intSetTsuki = Integer.Parse(.PropStrKikanTsuki) + (intTsukiCnt - 1)
                            If intSetTsuki > 12 Then
                                intSetTsuki -= 12
                            End If
                            .PropvwRiyoJokyoStudio.Sheets(0).ColumnHeader.Cells(0, intTsukiCnt - 1).Value = intSetTsuki.ToString & "月"
                            .PropvwRiyoJokyoStudio.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Studio_RIYO_KANO_NISSU, intTsukiCnt - 1).Value = dtRiyoJokyo.Rows(intDataRow)("RIYO_KANO_NISSU").ToString
                            intGokKano += Integer.Parse(dtRiyoJokyo.Rows(intDataRow)("RIYO_KANO_NISSU"))
                            .PropvwRiyoJokyoStudio.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Studio_KYUKAN_NISSU, intTsukiCnt - 1).Value = dtRiyoJokyo.Rows(intDataRow)("KYUKAN_NISSU").ToString
                            intGokKyukan += Integer.Parse(dtRiyoJokyo.Rows(intDataRow)("KYUKAN_NISSU"))
                            .PropvwRiyoJokyoStudio.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Studio_MAINTE_NISSU, intTsukiCnt - 1).Value = dtRiyoJokyo.Rows(intDataRow)("MAINTE_NISSU").ToString
                            intGokMainte += Integer.Parse(dtRiyoJokyo.Rows(intDataRow)("MAINTE_NISSU"))

                            .PropvwRiyoJokyoStudio.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Studio_RIYO_NISSU_201ST, intTsukiCnt - 1).Value = dtRiyoJokyo.Rows(intDataRow)("RIYO_NISSU_201st").ToString
                            intTotal += Integer.Parse(dtRiyoJokyo.Rows(intDataRow)("RIYO_NISSU_201st"))
                            intGok201st += Integer.Parse(dtRiyoJokyo.Rows(intDataRow)("RIYO_NISSU_201st"))
                            .PropvwRiyoJokyoStudio.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Studio_RIYO_NISSU_202ST, intTsukiCnt - 1).Value = dtRiyoJokyo.Rows(intDataRow)("RIYO_NISSU_202st").ToString
                            intTotal += Integer.Parse(dtRiyoJokyo.Rows(intDataRow)("RIYO_NISSU_202st"))
                            intGok202st += Integer.Parse(dtRiyoJokyo.Rows(intDataRow)("RIYO_NISSU_202st"))
                            .PropvwRiyoJokyoStudio.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Studio_RIYO_NISSU_HOUSELOCK, intTsukiCnt - 1).Value = dtRiyoJokyo.Rows(intDataRow)("RIYO_NISSU_houselock").ToString
                            intTotal += Integer.Parse(dtRiyoJokyo.Rows(intDataRow)("RIYO_NISSU_houselock"))
                            intGokhouselock += Integer.Parse(dtRiyoJokyo.Rows(intDataRow)("RIYO_NISSU_houselock"))
                            .PropvwRiyoJokyoStudio.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Studio_RIYO_NISSU_TOTAL, intTsukiCnt - 1).Value = intTotal.ToString
                            intGokTotal += intTotal

                            If Integer.Parse(dtRiyoJokyo.Rows(intDataRow)("RIYO_KANO_NISSU")) <> 0 Then
                                .PropvwRiyoJokyoStudio.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Studio_RIYO_RITU_TOTAL, intTsukiCnt - 1).Value = _
                                Math.Round(intTotal * 100 / Integer.Parse(dtRiyoJokyo.Rows(intDataRow)("RIYO_KANO_NISSU")), 0)
                            Else
                                .PropvwRiyoJokyoStudio.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Studio_RIYO_RITU_TOTAL, intTsukiCnt - 1).Value = "0"
                            End If
                            If intTotal <> 0 Then
                                .PropvwRiyoJokyoStudio.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Studio_RIYO_RITU_201ST, intTsukiCnt - 1).Value = _
                                Math.Round(Integer.Parse(dtRiyoJokyo.Rows(intDataRow)("RIYO_NISSU_201st")) * 100 / intTotal, 0)
                                .PropvwRiyoJokyoStudio.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Studio_RIYO_RITU_202ST, intTsukiCnt - 1).Value = _
                                Math.Round(Integer.Parse(dtRiyoJokyo.Rows(intDataRow)("RIYO_NISSU_202st")) * 100 / intTotal, 0)
                                .PropvwRiyoJokyoStudio.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Studio_RIYO_RITU_HOUSELOCK, intTsukiCnt - 1).Value = _
                                Math.Round(Integer.Parse(dtRiyoJokyo.Rows(intDataRow)("RIYO_NISSU_houselock")) * 100 / intTotal, 0)
                            Else
                                .PropvwRiyoJokyoStudio.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Studio_RIYO_RITU_201ST, intTsukiCnt - 1).Value = "0"
                                .PropvwRiyoJokyoStudio.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Studio_RIYO_RITU_202ST, intTsukiCnt - 1).Value = "0"
                                .PropvwRiyoJokyoStudio.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Studio_RIYO_RITU_HOUSELOCK, intTsukiCnt - 1).Value = "0"
                            End If
                        Next

                        If intTsukiCnt = 12 Then
                            '合計列表示
                            .PropvwRiyoJokyoStudio.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Studio_RIYO_KANO_NISSU, 12).Value = intGokKano.ToString
                            .PropvwRiyoJokyoStudio.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Studio_KYUKAN_NISSU, 12).Value = intGokKyukan.ToString
                            .PropvwRiyoJokyoStudio.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Studio_MAINTE_NISSU, 12).Value = intGokMainte.ToString
                            .PropvwRiyoJokyoStudio.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Studio_RIYO_NISSU_TOTAL, 12).Value = intGokTotal.ToString

                            If intGokKano <> 0 Then
                                .PropvwRiyoJokyoStudio.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Studio_RIYO_RITU_TOTAL, 12).Value = Math.Round(intGokTotal * 100 / intGokKano, 0).ToString
                            Else
                                .PropvwRiyoJokyoStudio.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Studio_RIYO_RITU_TOTAL, 12).Value = "0"
                            End If

                            .PropvwRiyoJokyoStudio.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Studio_RIYO_NISSU_201ST, 12).Value = intGok201st.ToString
                            .PropvwRiyoJokyoStudio.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Studio_RIYO_NISSU_202ST, 12).Value = intGok202st.ToString
                            .PropvwRiyoJokyoStudio.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Studio_RIYO_NISSU_HOUSELOCK, 12).Value = intGokhouselock.ToString

                            If intGokTotal <> 0 Then
                                .PropvwRiyoJokyoStudio.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Studio_RIYO_RITU_201ST, 12).Value = Math.Round(intGok201st * 100 / intGokTotal, 0).ToString
                                .PropvwRiyoJokyoStudio.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Studio_RIYO_RITU_202ST, 12).Value = Math.Round(intGok202st * 100 / intGokTotal, 0).ToString
                                .PropvwRiyoJokyoStudio.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Studio_RIYO_RITU_HOUSELOCK, 12).Value = Math.Round(intGokhouselock * 100 / intGokTotal, 0).ToString
                            Else
                                .PropvwRiyoJokyoStudio.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Studio_RIYO_RITU_201ST, 12).Value = "0"
                                .PropvwRiyoJokyoStudio.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Studio_RIYO_RITU_202ST, 12).Value = "0"
                                .PropvwRiyoJokyoStudio.Sheets(0).Cells(SpreadDayRiyoJokyoIndex_Studio_RIYO_RITU_HOUSELOCK, 12).Value = "0"
                            End If

                        End If

                    End If

                Next

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
    ''' <param name="dataEXTZ0104">データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了　false 異常終了</returns>
    ''' <remarks>スプレッドへのデータを取得する。
    ''' <para>作成情報：2015/09/24 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Private Function GetSpreadData(ByRef dataEXTZ0104 As DataEXTZ0104) As Boolean

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)        'コネクション
        Dim Adapter As New NpgsqlDataAdapter            'アダプタ
        Dim intNen As Integer = 0                       '年（データ取得用）
        Dim intTsuki As Integer = 0                     '月（データ取得用）
        Dim dtRiyoJokyo As DataTable = Nothing          '月毎のデータ

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try

            Cn.Open()

            '■日別売上データ取得
            'SELECT用SQLCommandを作成

            For intTsukiCnt As Integer = 1 To 12
                If intTsukiCnt > 1 Then
                    intTsuki += 1
                    If intTsuki > 12 Then
                        intNen += 1
                        intTsuki = 1
                    End If
                Else
                    intNen = Integer.Parse(dataEXTZ0104.PropStrKikanNen)
                    intTsuki = Integer.Parse(dataEXTZ0104.PropStrKikanTsuki)
                End If
                If sqlEXTZ0104.setSelectRiyoJokyo(Adapter, Cn, dataEXTZ0104.PropStrShisetsuKbn, intNen.ToString, intTsuki.ToString("00")) = False Then
                    '異常終了
                    Return False
                End If

                'ログ出力
                CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "", Nothing, Adapter.SelectCommand)

                'データ取得
                dtRiyoJokyo = New DataTable
                Adapter.Fill(dtRiyoJokyo)
                Select Case intTsukiCnt
                    Case 1
                        dataEXTZ0104.PropDtRiyoJokyo1 = dtRiyoJokyo
                    Case 2
                        dataEXTZ0104.PropDtRiyoJokyo2 = dtRiyoJokyo
                    Case 3
                        dataEXTZ0104.PropDtRiyoJokyo3 = dtRiyoJokyo
                    Case 4
                        dataEXTZ0104.PropDtRiyoJokyo4 = dtRiyoJokyo
                    Case 5
                        dataEXTZ0104.PropDtRiyoJokyo5 = dtRiyoJokyo
                    Case 6
                        dataEXTZ0104.PropDtRiyoJokyo6 = dtRiyoJokyo
                    Case 7
                        dataEXTZ0104.PropDtRiyoJokyo7 = dtRiyoJokyo
                    Case 8
                        dataEXTZ0104.PropDtRiyoJokyo8 = dtRiyoJokyo
                    Case 9
                        dataEXTZ0104.PropDtRiyoJokyo9 = dtRiyoJokyo
                    Case 10
                        dataEXTZ0104.PropDtRiyoJokyo10 = dtRiyoJokyo
                    Case 11
                        dataEXTZ0104.PropDtRiyoJokyo11 = dtRiyoJokyo
                    Case 12
                        dataEXTZ0104.PropDtRiyoJokyo12 = dtRiyoJokyo
                End Select
            Next

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
            dtRiyoJokyo.Dispose()

        End Try

    End Function

#End Region

End Class
