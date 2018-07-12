Imports Common
Imports CommonEXT
Imports Npgsql

''' <summary>
''' 請求一覧で使用する情報取得SQL作成
''' </summary>
''' <remarks>請求一覧で使用する情報取得のSQLを作成する
''' <para>作成情報：2015/09/15 h.endo
''' <p>改訂情報:</p>
''' </para></remarks>
Public Class SqlEXTZ0102

    '*****基本SQL文*****

#Region "SQL文(請求一覧（シアター）取得)"
    'SQL文(請求一覧（シアター）取得)
    ' 2016.01.13 MOD START↓ y.morooka 入金情報複数件の対応 入金額、入金状況の表示追加
    '    Private strSelectSeikyuTheatreSQL As String = _
    '"SELECT" & vbCrLf & _
    '"    t1.SEIKYU_IRAI_NO," & vbCrLf & _
    '"    t1.SEIKYU_DT," & vbCrLf & _
    '"    t2a.YOYAKU_DT_FROM," & vbCrLf & _
    '"    t2a.YOYAKU_DT_TO," & vbCrLf & _
    '"    t2.SAIJI_NM," & vbCrLf & _
    '"    t3.AITE_NM," & vbCrLf & _
    '"    t1.SEIKYU_KIN," & vbCrLf & _
    '"    CASE t1.SEIKYU_NAIYO" & vbCrLf & _
    '"        WHEN '1' THEN '利用料'" & vbCrLf & _
    '"        WHEN '2' THEN '付帯設備'" & vbCrLf & _
    '"        WHEN '3' THEN '利用料＋付帯設備'" & vbCrLf & _
    '"        WHEN '4' THEN '還付'" & vbCrLf & _
    '"    END SEIKYU_NAIYO," & vbCrLf & _
    '"    t1.NYUKIN_YOTEI_DT," & vbCrLf & _
    '"    t1.NYUKIN_DT," & vbCrLf & _
    '"    t1.YOYAKU_NO" & vbCrLf & _
    '"FROM" & vbCrLf & _
    '"    BILLPAY_TBL t1" & vbCrLf & _
    '"    LEFT JOIN YOYAKU_TBL t2" & vbCrLf & _
    '"    ON t1.YOYAKU_NO = t2.YOYAKU_NO" & vbCrLf & _
    '"    LEFT JOIN" & vbCrLf & _
    '"        (" & vbCrLf & _
    '"            SELECT" & vbCrLf & _
    '"            	YOYAKU_NO," & vbCrLf & _
    '"                MIN(YOYAKU_DT) AS YOYAKU_DT_FROM," & vbCrLf & _
    '"                MAX(YOYAKU_DT) AS YOYAKU_DT_TO" & vbCrLf & _
    '"            FROM" & vbCrLf & _
    '"                YDT_TBL" & vbCrLf & _
    '"            WHERE" & vbCrLf & _
    '"                SHISETU_KBN = '1'" & vbCrLf & _
    '"            GROUP BY YOYAKU_NO" & vbCrLf & _
    '"        ) t2a" & vbCrLf & _
    '"    ON  t1.YOYAKU_NO = t2a.YOYAKU_NO" & vbCrLf & _
    '"    LEFT JOIN AITESAKI_MST t3" & vbCrLf & _
    '"    ON  t1.AITE_CD = t3.AITE_CD" & vbCrLf & _
    '"WHERE" & vbCrLf & _
    '"    t1.SEIKYU_IRAI_FLG = '1'" & vbCrLf & _
    '"AND t2.SHISETU_KBN = '1'" & vbCrLf

    Private strSelectSeikyuTheatreSQL As String = _
"SELECT" & vbCrLf & _
"    t1.SEIKYU_IRAI_NO," & vbCrLf & _
"    t1.SEIKYU_DT," & vbCrLf & _
"    t2a.YOYAKU_DT_FROM," & vbCrLf & _
"    t2a.YOYAKU_DT_TO," & vbCrLf & _
"    t2.SAIJI_NM," & vbCrLf & _
"    t3.AITE_NM," & vbCrLf & _
"    t1.SEIKYU_KIN," & vbCrLf & _
"    CASE t1.SEIKYU_NAIYO" & vbCrLf & _
"        WHEN '1' THEN '利用料'" & vbCrLf & _
"        WHEN '2' THEN '付帯設備'" & vbCrLf & _
"        WHEN '3' THEN '利用料＋付帯設備'" & vbCrLf & _
"        WHEN '4' THEN '還付'" & vbCrLf & _
"    END SEIKYU_NAIYO," & vbCrLf & _
"    t1.NYUKIN_YOTEI_DT," & vbCrLf & _
"    t1.NYUKIN_DT," & vbCrLf & _
"    t1.NYUKIN_KIN," & vbCrLf & _
"    CASE t1.NYUKIN_INPUT_FLG" & vbCrLf & _
"        WHEN '0' THEN ''" & vbCrLf & _
"        WHEN '1' THEN " & vbCrLf & _
"           CASE" & vbCrLf & _
"               WHEN t1.SEIKYU_KIN <> t1.NYUKIN_KIN THEN ''" & vbCrLf & _
"               WHEN t1.SEIKYU_KIN = t1.NYUKIN_KIN THEN '○'" & vbCrLf & _
"           END" & vbCrLf & _
"    END NYUKIN_INPUT_FLG," & vbCrLf & _
"    t1.YOYAKU_NO" & vbCrLf & _
"FROM" & vbCrLf & _
"    BILLPAY_TBL t1" & vbCrLf & _
"    LEFT JOIN YOYAKU_TBL t2" & vbCrLf & _
"    ON t1.YOYAKU_NO = t2.YOYAKU_NO" & vbCrLf & _
"    LEFT JOIN" & vbCrLf & _
"        (" & vbCrLf & _
"            SELECT" & vbCrLf & _
"            	YOYAKU_NO," & vbCrLf & _
"                MIN(YOYAKU_DT) AS YOYAKU_DT_FROM," & vbCrLf & _
"                MAX(YOYAKU_DT) AS YOYAKU_DT_TO" & vbCrLf & _
"            FROM" & vbCrLf & _
"                YDT_TBL" & vbCrLf & _
"            WHERE" & vbCrLf & _
"                SHISETU_KBN = '1'" & vbCrLf & _
"            GROUP BY YOYAKU_NO" & vbCrLf & _
"        ) t2a" & vbCrLf & _
"    ON  t1.YOYAKU_NO = t2a.YOYAKU_NO" & vbCrLf & _
"    LEFT JOIN AITESAKI_MST t3" & vbCrLf & _
"    ON  t1.AITE_CD = t3.AITE_CD" & vbCrLf & _
"WHERE" & vbCrLf & _
"    t1.SEIKYU_IRAI_FLG = '1'" & vbCrLf & _
"AND t2.SHISETU_KBN = '1'" & vbCrLf
    ' 2016.01.13 MOD END↑ y.morooka 入金情報複数件の対応 入金額、入金状況の表示追加
#End Region

#Region "SQL文(請求一覧（スタジオ）取得)"

    'SQL文(請求一覧（スタジオ）取得)
    ' 2016.01.13 MOD START↓ y.morooka 入金情報複数件の対応 入金額、入金状況の表示追加
    '    Private strSelectSeikyuStudioSQL As String = _
    '"SELECT" & vbCrLf & _
    '"    t1.SEIKYU_IRAI_NO," & vbCrLf & _
    '"    t1.SEIKYU_DT," & vbCrLf & _
    '"    CASE t2.STUDIO_KBN" & vbCrLf & _
    '"    WHEN '1' THEN '201st'" & vbCrLf & _
    '"    WHEN '2' THEN '202st'" & vbCrLf & _
    '"    WHEN '3' THEN 'house lock'" & vbCrLf & _
    '"    END STUDIO," & vbCrLf & _
    '"    t2a.YOYAKU_DT_FROM," & vbCrLf & _
    '"    t2a.YOYAKU_DT_TO," & vbCrLf & _
    '"    t2.SHUTSUEN_NM," & vbCrLf & _
    '"    t3.AITE_NM," & vbCrLf & _
    '"    t1.SEIKYU_KIN," & vbCrLf & _
    '"    CASE t1.SEIKYU_NAIYO" & vbCrLf & _
    '"        WHEN '1' THEN '利用料'" & vbCrLf & _
    '"        WHEN '2' THEN '付帯設備'" & vbCrLf & _
    '"        WHEN '3' THEN '利用料＋付帯設備'" & vbCrLf & _
    '"        WHEN '4' THEN '還付'" & vbCrLf & _
    '"    END SEIKYU_NAIYO," & vbCrLf & _
    '"    t1.NYUKIN_YOTEI_DT," & vbCrLf & _
    '"    t1.NYUKIN_DT," & vbCrLf & _
    '"    t1.YOYAKU_NO" & vbCrLf & _
    '"FROM" & vbCrLf & _
    '"    BILLPAY_TBL t1" & vbCrLf & _
    '"    LEFT JOIN YOYAKU_TBL t2" & vbCrLf & _
    '"    ON t1.YOYAKU_NO = t2.YOYAKU_NO" & vbCrLf & _
    '"    LEFT JOIN" & vbCrLf & _
    '"        (" & vbCrLf & _
    '"            SELECT" & vbCrLf & _
    '"            	YOYAKU_NO," & vbCrLf & _
    '"                MIN(YOYAKU_DT) AS YOYAKU_DT_FROM," & vbCrLf & _
    '"                MAX(YOYAKU_DT) AS YOYAKU_DT_TO" & vbCrLf & _
    '"            FROM" & vbCrLf & _
    '"                YDT_TBL" & vbCrLf & _
    '"            WHERE" & vbCrLf & _
    '"                SHISETU_KBN = '2'" & vbCrLf & _
    '"            GROUP BY YOYAKU_NO" & vbCrLf & _
    '"        ) t2a" & vbCrLf & _
    '"    ON  t1.YOYAKU_NO = T2a.YOYAKU_NO" & vbCrLf & _
    '"    LEFT JOIN AITESAKI_MST t3" & vbCrLf & _
    '"    ON  t1.AITE_CD = t3.AITE_CD" & vbCrLf & _
    '"WHERE" & vbCrLf & _
    '"    t1.SEIKYU_IRAI_FLG = '1'" & vbCrLf & _
    '"AND t2.SHISETU_KBN = '2'" & vbCrLf
    Private strSelectSeikyuStudioSQL As String = _
"SELECT" & vbCrLf & _
"    t1.SEIKYU_IRAI_NO," & vbCrLf & _
"    t1.SEIKYU_DT," & vbCrLf & _
"    CASE t2.STUDIO_KBN" & vbCrLf & _
"    WHEN '1' THEN '201st'" & vbCrLf & _
"    WHEN '2' THEN '202st'" & vbCrLf & _
"    WHEN '3' THEN 'house lock'" & vbCrLf & _
"    END STUDIO," & vbCrLf & _
"    t2a.YOYAKU_DT_FROM," & vbCrLf & _
"    t2a.YOYAKU_DT_TO," & vbCrLf & _
"    t2.SHUTSUEN_NM," & vbCrLf & _
"    t3.AITE_NM," & vbCrLf & _
"    t1.SEIKYU_KIN," & vbCrLf & _
"    CASE t1.SEIKYU_NAIYO" & vbCrLf & _
"        WHEN '1' THEN '利用料'" & vbCrLf & _
"        WHEN '2' THEN '付帯設備'" & vbCrLf & _
"        WHEN '3' THEN '利用料＋付帯設備'" & vbCrLf & _
"        WHEN '4' THEN '還付'" & vbCrLf & _
"    END SEIKYU_NAIYO," & vbCrLf & _
"    t1.NYUKIN_YOTEI_DT," & vbCrLf & _
"    t1.NYUKIN_DT," & vbCrLf & _
"    t1.NYUKIN_KIN," & vbCrLf & _
"    CASE t1.NYUKIN_INPUT_FLG" & vbCrLf & _
"        WHEN '0' THEN ''" & vbCrLf & _
"        WHEN '1' THEN " & vbCrLf & _
"           CASE" & vbCrLf & _
"               WHEN t1.SEIKYU_KIN <> t1.NYUKIN_KIN THEN ''" & vbCrLf & _
"               WHEN t1.SEIKYU_KIN = t1.NYUKIN_KIN THEN '○'" & vbCrLf & _
"           END" & vbCrLf & _
"    END NYUKIN_INPUT_FLG," & vbCrLf & _
"    t1.YOYAKU_NO" & vbCrLf & _
"FROM" & vbCrLf & _
"    BILLPAY_TBL t1" & vbCrLf & _
"    LEFT JOIN YOYAKU_TBL t2" & vbCrLf & _
"    ON t1.YOYAKU_NO = t2.YOYAKU_NO" & vbCrLf & _
"    LEFT JOIN" & vbCrLf & _
"        (" & vbCrLf & _
"            SELECT" & vbCrLf & _
"            	YOYAKU_NO," & vbCrLf & _
"                MIN(YOYAKU_DT) AS YOYAKU_DT_FROM," & vbCrLf & _
"                MAX(YOYAKU_DT) AS YOYAKU_DT_TO" & vbCrLf & _
"            FROM" & vbCrLf & _
"                YDT_TBL" & vbCrLf & _
"            WHERE" & vbCrLf & _
"                SHISETU_KBN = '2'" & vbCrLf & _
"            GROUP BY YOYAKU_NO" & vbCrLf & _
"        ) t2a" & vbCrLf & _
"    ON  t1.YOYAKU_NO = T2a.YOYAKU_NO" & vbCrLf & _
"    LEFT JOIN AITESAKI_MST t3" & vbCrLf & _
"    ON  t1.AITE_CD = t3.AITE_CD" & vbCrLf & _
"WHERE" & vbCrLf & _
"    t1.SEIKYU_IRAI_FLG = '1'" & vbCrLf & _
"AND t2.SHISETU_KBN = '2'" & vbCrLf
    ' 2016.01.13 MOD END↑ y.morooka 入金情報複数件の対応 入金額、入金状況の表示追加

#End Region

    '*****条件追加・変更、アダプタ・パラメータ網羅によるSQL文*****
#Region "請求一覧取得"

    ''' <summary>
    ''' 請求一覧取得
    ''' </summary>
    ''' <param name="Adapter">アダプタ</param>
    ''' <param name="Cn">コネクション</param>
    ''' <param name="dataEXTZ0102">データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>請求一覧を取得するSQLの作成
    ''' <para>作成情報：2015/09/15 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function setSelectSeikyu(ByRef Adapter As NpgsqlDataAdapter, _
                                    ByRef Cn As NpgsqlConnection, _
                                    ByVal dataEXTZ0102 As DataEXTZ0102) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = String.Empty

        Try

            With dataEXTZ0102
                'SQL文(SELECT)
                'データアダプタに、SQLのSELECT文を設定
                If .PropStrShisetsuKbn = SHISETU_KBN_THEATER Then
                    'シアターの場合
                    strSQL = strSelectSeikyuTheatreSQL
                Else
                    'スタジオの場合
                    strSQL = strSelectSeikyuStudioSQL
                End If

                '請求日(From）が入力されている場合
                If .PropStrSeikyuDtFrom <> String.Empty Then
                    strSQL &= "AND TO_DATE(t1.SEIKYU_DT, 'YYYY/MM/DD') >= TO_DATE('" & .PropStrSeikyuDtFrom & "', 'YYYY/MM/DD') " & vbCrLf
                End If

                '請求日(To）が入力されている場合
                If .PropStrSeikyuDtTo <> String.Empty Then
                    strSQL &= "AND TO_DATE(t1.SEIKYU_DT, 'YYYY/MM/DD') <= TO_DATE('" & .PropStrSeikyuDtTo & "', 'YYYY/MM/DD') " & vbCrLf
                End If

                '入金予定日(From）が入力されている場合
                If .PropStrNyukinYoteiDtFrom <> String.Empty Then
                    strSQL &= "AND TO_DATE(t1.NYUKIN_YOTEI_DT, 'YYYY/MM/DD') >= TO_DATE('" & .PropStrNyukinYoteiDtFrom & "', 'YYYY/MM/DD') " & vbCrLf
                End If

                '入金予定日(To）が入力されている場合
                If .PropStrNyukinYoteiDtTo <> String.Empty Then
                    strSQL &= "AND TO_DATE(t1.NYUKIN_YOTEI_DT, 'YYYY/MM/DD') <= TO_DATE('" & .PropStrNyukinYoteiDtTo & "', 'YYYY/MM/DD') " & vbCrLf
                End If

                '入金日(From）が入力されている場合
                If .PropStrNyukinDtFrom <> String.Empty Then
                    strSQL &= "AND TO_DATE(t1.NYUKIN_DT, 'YYYY/MM/DD') >= TO_DATE('" & .PropStrNyukinDtFrom & "', 'YYYY/MM/DD') " & vbCrLf
                End If

                '入金日(To）が入力されている場合
                If .PropStrNyukinDtTo <> String.Empty Then
                    strSQL &= "AND TO_DATE(t1.NYUKIN_DT, 'YYYY/MM/DD') <= TO_DATE('" & .PropStrNyukinDtTo & "', 'YYYY/MM/DD') " & vbCrLf
                End If

                '相手先名が入力されている場合
                If .PropStrAiteNm <> String.Empty Then
                    strSQL &= "AND t3.AITE_NM LIKE '%' || '" & .PropStrAiteNm & "' || '%'" & vbCrLf
                End If

                '相手先名カナが入力されている場合
                If .PropStrAiteNmKana <> String.Empty Then
                    strSQL &= "AND t3.AITE_NM_KANA LIKE '%' || '" & .PropStrAiteNmKana & "' || '%'" & vbCrLf
                End If

                '利用者名が入力されている場合
                If .PropStrRiyoNm <> String.Empty Then
                    strSQL &= "AND t2.RIYO_NM LIKE '%' || '" & .PropStrRiyoNm & "' || '%'" & vbCrLf
                End If

                '催事名が入力されている場合
                If .PropStrSaijiNm <> String.Empty Then
                    If .PropStrShisetsuKbn = SHISETU_KBN_THEATER Then
                        'シアターの場合
                        strSQL &= "AND t2.SAIJI_NM LIKE '%' || '" & .PropStrSaijiNm & "' || '%'" & vbCrLf
                    Else
                        'スタジオの場合
                        strSQL &= "AND t2.SHUTSUEN_NM LIKE '%' || '" & .PropStrSaijiNm & "' || '%'" & vbCrLf
                    End If
                End If

                '請求依頼番号が入力されている場合
                If .PropStrSeikyuIraiNo <> String.Empty Then
                    strSQL &= "AND t1.SEIKYU_IRAI_NO = '" & .PropStrSeikyuIraiNo & "' " & vbCrLf
                End If

                '未入金に関する出力判定
                If .PropBlnMinyukin Then
                    '入金予定日を過ぎた未入金請求のみを抽出
                    strSQL &= "AND (TO_DATE(t1.NYUKIN_YOTEI_DT,('YYYY/MM/DD')) < CURRENT_DATE AND t1.NYUKIN_INPUT_FLG = '0')"
                End If
                ' 2015.12.04 ADD START↓ h.hagiwara ソート順設定
                strSQL &= " ORDER BY SEIKYU_DT DESC ,  YOYAKU_NO "
                ' 2015.12.04 ADD END↑ h.hagiwara ソート順設定

            End With

            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)

            '例外処理
            Return False
        End Try
    End Function

#End Region


End Class
