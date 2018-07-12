Imports Common
Imports CommonEXT
Imports Npgsql

''' <summary>
''' 予約一覧で使用する情報取得SQL作成
''' </summary>
''' <remarks>予約一覧で使用する情報取得のSQLを作成する
''' <para>作成情報：2015/09/09 h.endo
''' <p>改訂情報:</p>
''' </para></remarks>
Public Class SqlEXTZ0101

    '*****基本SQL文*****

#Region "SQL文(予約一覧（シアター）取得)"

    'SQL文(予約一覧（シアター）取得)
    ' 2016.1.15 START ↓y.ozawa
    'Private strSelectYoyakuTheatreSQL As String = _
    '"SELECT" & vbCrLf & _
    '    "DISTINCT" & vbCrLf & _
    '    "t1.YOYAKU_NO," & vbCrLf & _
    '    "CASE t1.YOYAKU_STS" & vbCrLf & _
    '    "    WHEN '1' THEN '仮未'" & vbCrLf & _
    '    "    WHEN '2' THEN '仮'" & vbCrLf & _
    '    "    WHEN '3' THEN '正式'" & vbCrLf & _
    '    "    WHEN '4' THEN '完了'" & vbCrLf & _
    '    "END YOYAKU_STS," & vbCrLf & _
    '    "t2a.YOYAKU_DT_FROM," & vbCrLf & _
    '    "t2a.YOYAKU_DT_TO," & vbCrLf & _
    '    "t2b.START_TIME," & vbCrLf & _
    '    "t2c.END_TIME," & vbCrLf & _
    '    "t1.SAIJI_NM," & vbCrLf & _
    '    "CASE t1.RIYO_TYPE" & vbCrLf & _
    '    "    WHEN '0' THEN '未定'" & vbCrLf & _
    '    "    WHEN '1' THEN 'スタンディング'" & vbCrLf & _
    '    "    WHEN '2' THEN '着席'" & vbCrLf & _
    '    "    WHEN '3' THEN '変則'" & vbCrLf & _
    '    "    WHEN '4' THEN '催事'" & vbCrLf & _
    '    "END RIYO_TYPE," & vbCrLf & _
    '    "t1.RIYO_NM," & vbCrLf & _
    '    "t1.SEKININ_NM," & vbCrLf & _
    '    "t3.SHONIN_NINZU," & vbCrLf & _
    '    "CASE t4a.RIYO_KIN" & vbCrLf & _
    '    "    WHEN '0' THEN ''" & vbCrLf & _
    '    "    WHEN '1' THEN '△'" & vbCrLf & _
    '    "    WHEN '2' THEN '〇'" & vbCrLf & _
    '    "    WHEN '3' THEN '◎'" & vbCrLf & _
    '    "END RIYO_KIN," & vbCrLf & _
    '    "CASE" & vbCrLf & _
    '    "    WHEN t1.FINPUT_STS = '1' THEN '確定'" & vbCrLf & _
    '    "END FINPUT_STS," & vbCrLf & _
    '    "CASE t4b.FUTAI_SETUBI" & vbCrLf & _
    '    "    WHEN '0' THEN ''" & vbCrLf & _
    '    "    WHEN '1' THEN '△'" & vbCrLf & _
    '    "    WHEN '2' THEN '〇'" & vbCrLf & _
    '    "    WHEN '3' THEN '◎'" & vbCrLf & _
    '    "END FUTAI_SETUBI," & vbCrLf & _
    '    "t1.YOYAKU_STS AS STS_CD" & vbCrLf & _
    '"FROM" & vbCrLf & _
    '   " YOYAKU_TBL t1" & vbCrLf & _
    '   " LEFT JOIN YDT_TBL t2" & vbCrLf & _
    '   " ON  t1.YOYAKU_NO = T2.YOYAKU_NO" & vbCrLf & _
    '   " LEFT JOIN" & vbCrLf & _
    '   "     (" & vbCrLf & _
    '   "         SELECT" & vbCrLf & _
    '   "             YOYAKU_NO," & vbCrLf & _
    '   "             MIN(YOYAKU_DT) AS YOYAKU_DT_FROM," & vbCrLf & _
    '   "             MAX(YOYAKU_DT) AS YOYAKU_DT_TO" & vbCrLf & _
    '   "         FROM" & vbCrLf & _
    '   "             YDT_TBL" & vbCrLf & _
    '   "         WHERE" & vbCrLf & _
    '   "             SHISETU_KBN = '1'" & vbCrLf & _
    '   "         GROUP BY YOYAKU_NO" & vbCrLf & _
    '   "     ) t2a" & vbCrLf & _
    '   " ON  t1.YOYAKU_NO = T2a.YOYAKU_NO" & vbCrLf & _
    '   " LEFT JOIN" & vbCrLf & _
    '   "     (   SELECT" & vbCrLf & _
    '   "             YOYAKU_NO," & vbCrLf & _
    '   "             YOYAKU_DT," & vbCrLf & _
    '   "             START_TIME" & vbCrLf & _
    '   "         FROM" & vbCrLf & _
    '   "             YDT_TBL" & vbCrLf & _
    '   "         WHERE" & vbCrLf & _
    '   "             SHISETU_KBN = '1'" & vbCrLf & _
    '   "     ) t2b" & vbCrLf & _
    '   " ON  t2a.YOYAKU_NO = T2b.YOYAKU_NO" & vbCrLf & _
    '   " AND t2a.YOYAKU_DT_FROM = T2b.YOYAKU_DT" & vbCrLf & _
    '   " LEFT JOIN" & vbCrLf & _
    '   "     (" & vbCrLf & _
    '   "         SELECT" & vbCrLf & _
    '   "             YOYAKU_NO," & vbCrLf & _
    '   "             YOYAKU_DT," & vbCrLf & _
    '   "             END_TIME" & vbCrLf & _
    '   "         FROM" & vbCrLf & _
    '   "             YDT_TBL" & vbCrLf & _
    '   "         WHERE" & vbCrLf & _
    '   "             SHISETU_KBN = '1'" & vbCrLf & _
    '   "     ) t2c" & vbCrLf & _
    '   " ON  t2a.YOYAKU_NO = T2c.YOYAKU_NO" & vbCrLf & _
    '   " AND t2a.YOYAKU_DT_TO = T2c.YOYAKU_DT" & vbCrLf & _
    '   " LEFT JOIN" & vbCrLf & _
    '   "     (" & vbCrLf & _
    '   "         SELECT" & vbCrLf & _
    '   "             YOYAKU_NO," & vbCrLf & _
    '   "             COUNT(YOYAKU_NO) AS SHONIN_NINZU" & vbCrLf & _
    '   "         FROM" & vbCrLf & _
    '   "             CHECK_RIREKI_TBL" & vbCrLf & _
    '   "         WHERE" & vbCrLf & _
    '   "             CHECK_STS = '1'" & vbCrLf & _
    '   "         GROUP BY" & vbCrLf & _
    '   "             YOYAKU_NO" & vbCrLf & _
    '   "     ) t3" & vbCrLf & _
    '   " ON  t1.YOYAKU_NO = T3.YOYAKU_NO" & vbCrLf & _
    '   " LEFT JOIN" & vbCrLf & _
    '   "     (" & vbCrLf & _
    '   "         SELECT" & vbCrLf & _
    '   "             YOYAKU_NO," & vbCrLf & _
    '   "             MIN(tmp.RIYO_KIN) RIYO_KIN" & vbCrLf & _
    '   "         FROM" & vbCrLf & _
    '   "             (" & vbCrLf & _
    '   "                 SELECT" & vbCrLf & _
    '   "                     YOYAKU_NO," & vbCrLf & _
    '   "                     CASE" & vbCrLf & _
    '   "                         WHEN SEIKYU_INPUT_FLG = '0' THEN '0'" & vbCrLf & _
    '   "                         ELSE CASE" & vbCrLf & _
    '   "                             WHEN SEIKYU_IRAI_FLG = '0' THEN '1'" & vbCrLf & _
    '   "                             ELSE CASE" & vbCrLf & _
    '   "                                 WHEN NYUKIN_INPUT_FLG = '0' THEN '2'" & vbCrLf & _
    '   "                                 ELSE '3'" & vbCrLf & _
    '   "                             END" & vbCrLf & _
    '   "                         END" & vbCrLf & _
    '   "                     END RIYO_KIN" & vbCrLf & _
    '   "                 FROM" & vbCrLf & _
    '   "                     BILLPAY_TBL" & vbCrLf & _
    '   "                 WHERE" & vbCrLf & _
    '   "                     SEIKYU_NAIYO = '1'" & vbCrLf & _
    '   "                 OR  SEIKYU_NAIYO = '3'" & vbCrLf & _
    '   "             )tmp" & vbCrLf & _
    '   "        GROUP BY YOYAKU_NO" & vbCrLf & _
    '   "     ) t4a" & vbCrLf & _
    '   " ON  t1.YOYAKU_NO = t4a.YOYAKU_NO" & vbCrLf & _
    '   " LEFT JOIN" & vbCrLf & _
    '   "     (" & vbCrLf & _
    '   "         SELECT" & vbCrLf & _
    '   "             YOYAKU_NO," & vbCrLf & _
    '   "             MIN(tmp.FUTAI_SETUBI) FUTAI_SETUBI" & vbCrLf & _
    '   "         FROM" & vbCrLf & _
    '   "             (" & vbCrLf & _
    '   "                 SELECT" & vbCrLf & _
    '   "                     YOYAKU_NO," & vbCrLf & _
    '   "                     CASE" & vbCrLf & _
    '   "                         WHEN SEIKYU_INPUT_FLG = '0' THEN '0'" & vbCrLf & _
    '   "                         ELSE CASE" & vbCrLf & _
    '   "                             WHEN SEIKYU_IRAI_FLG = '0' THEN '1'" & vbCrLf & _
    '   "                             ELSE CASE" & vbCrLf & _
    '   "                                 WHEN NYUKIN_INPUT_FLG = '0' THEN '2'" & vbCrLf & _
    '   "                                 ELSE '3'" & vbCrLf & _
    '   "                             END" & vbCrLf & _
    '   "                         END" & vbCrLf & _
    '   "                     END FUTAI_SETUBI" & vbCrLf & _
    '   "                 FROM" & vbCrLf & _
    '   "                     BILLPAY_TBL" & vbCrLf & _
    '   "                 WHERE" & vbCrLf & _
    '   "                     SEIKYU_NAIYO = '2'" & vbCrLf & _
    '   "                 OR  SEIKYU_NAIYO = '3'" & vbCrLf & _
    '   "             )tmp" & vbCrLf & _
    '   "        GROUP BY YOYAKU_NO" & vbCrLf & _
    '   "     ) t4b" & vbCrLf & _
    '   " ON t1.YOYAKU_NO = t4b.YOYAKU_NO" & vbCrLf & _
    '   " LEFT JOIN BILLPAY_TBL t5" & vbCrLf & _
    '   " ON t1.YOYAKU_NO = t5.YOYAKU_NO" & vbCrLf & _
    '"WHERE" & vbCrLf & _
    '"    t1.SHISETU_KBN = '1'"
    ' 2016.04.04 UPD START↓ h.hagiwara 条件変更に伴い参照テーブル削除
    'Private strSelectYoyakuTheatreSQL As String = _
    '"SELECT" & vbCrLf & _
    '    "DISTINCT" & vbCrLf & _
    '    "t1.YOYAKU_NO," & vbCrLf & _
    '    "CASE t1.YOYAKU_STS" & vbCrLf & _
    '    "    WHEN '1' THEN '仮'" & vbCrLf & _
    '    "    WHEN '2' THEN '決定'" & vbCrLf & _
    '    "    WHEN '3' THEN '受諾'" & vbCrLf & _
    '    "    WHEN '4' THEN '完了'" & vbCrLf & _
    '    "END YOYAKU_STS," & vbCrLf & _
    '    "t2a.YOYAKU_DT_FROM," & vbCrLf & _
    '    "t2a.YOYAKU_DT_TO," & vbCrLf & _
    '    "t2b.START_TIME," & vbCrLf & _
    '    "t2c.END_TIME," & vbCrLf & _
    '    "t1.SAIJI_NM," & vbCrLf & _
    '    "CASE t1.RIYO_TYPE" & vbCrLf & _
    '    "    WHEN '0' THEN '未定'" & vbCrLf & _
    '    "    WHEN '1' THEN 'スタンディング'" & vbCrLf & _
    '    "    WHEN '2' THEN '着席'" & vbCrLf & _
    '    "    WHEN '3' THEN '変則'" & vbCrLf & _
    '    "    WHEN '4' THEN '催事'" & vbCrLf & _
    '    "END RIYO_TYPE," & vbCrLf & _
    '    "t1.RIYO_NM," & vbCrLf & _
    '    "t1.SEKININ_NM," & vbCrLf & _
    '    "t3.SHONIN_NINZU," & vbCrLf & _
    '    "CASE t4a.RIYO_KIN" & vbCrLf & _
    '    "    WHEN '0' THEN ''" & vbCrLf & _
    '    "    WHEN '1' THEN '△'" & vbCrLf & _
    '    "    WHEN '2' THEN '〇'" & vbCrLf & _
    '    "    WHEN '3' THEN '◎'" & vbCrLf & _
    '    "END RIYO_KIN," & vbCrLf & _
    '    "CASE" & vbCrLf & _
    '    "    WHEN t1.FINPUT_STS = '1' THEN '確定'" & vbCrLf & _
    '    "END FINPUT_STS," & vbCrLf & _
    '    "CASE t4b.FUTAI_SETUBI" & vbCrLf & _
    '    "    WHEN '0' THEN ''" & vbCrLf & _
    '    "    WHEN '1' THEN '△'" & vbCrLf & _
    '    "    WHEN '2' THEN '〇'" & vbCrLf & _
    '    "    WHEN '3' THEN '◎'" & vbCrLf & _
    '    "END FUTAI_SETUBI," & vbCrLf & _
    '    "t1.YOYAKU_STS AS STS_CD" & vbCrLf & _
    '"FROM" & vbCrLf & _
    '   " YOYAKU_TBL t1" & vbCrLf & _
    '   " LEFT JOIN YDT_TBL t2" & vbCrLf & _
    '   " ON  t1.YOYAKU_NO = T2.YOYAKU_NO" & vbCrLf & _
    '   " LEFT JOIN" & vbCrLf & _
    '   "     (" & vbCrLf & _
    '   "         SELECT" & vbCrLf & _
    '   "             YOYAKU_NO," & vbCrLf & _
    '   "             MIN(YOYAKU_DT) AS YOYAKU_DT_FROM," & vbCrLf & _
    '   "             MAX(YOYAKU_DT) AS YOYAKU_DT_TO" & vbCrLf & _
    '   "         FROM" & vbCrLf & _
    '   "             YDT_TBL" & vbCrLf & _
    '   "         WHERE" & vbCrLf & _
    '   "             SHISETU_KBN = '1'" & vbCrLf & _
    '   "         GROUP BY YOYAKU_NO" & vbCrLf & _
    '   "     ) t2a" & vbCrLf & _
    '   " ON  t1.YOYAKU_NO = T2a.YOYAKU_NO" & vbCrLf & _
    '   " LEFT JOIN" & vbCrLf & _
    '   "     (   SELECT" & vbCrLf & _
    '   "             YOYAKU_NO," & vbCrLf & _
    '   "             YOYAKU_DT," & vbCrLf & _
    '   "             START_TIME" & vbCrLf & _
    '   "         FROM" & vbCrLf & _
    '   "             YDT_TBL" & vbCrLf & _
    '   "         WHERE" & vbCrLf & _
    '   "             SHISETU_KBN = '1'" & vbCrLf & _
    '   "     ) t2b" & vbCrLf & _
    '   " ON  t2a.YOYAKU_NO = T2b.YOYAKU_NO" & vbCrLf & _
    '   " AND t2a.YOYAKU_DT_FROM = T2b.YOYAKU_DT" & vbCrLf & _
    '   " LEFT JOIN" & vbCrLf & _
    '   "     (" & vbCrLf & _
    '   "         SELECT" & vbCrLf & _
    '   "             YOYAKU_NO," & vbCrLf & _
    '   "             YOYAKU_DT," & vbCrLf & _
    '   "             END_TIME" & vbCrLf & _
    '   "         FROM" & vbCrLf & _
    '   "             YDT_TBL" & vbCrLf & _
    '   "         WHERE" & vbCrLf & _
    '   "             SHISETU_KBN = '1'" & vbCrLf & _
    '   "     ) t2c" & vbCrLf & _
    '   " ON  t2a.YOYAKU_NO = T2c.YOYAKU_NO" & vbCrLf & _
    '   " AND t2a.YOYAKU_DT_TO = T2c.YOYAKU_DT" & vbCrLf & _
    '   " LEFT JOIN" & vbCrLf & _
    '   "     (" & vbCrLf & _
    '   "         SELECT" & vbCrLf & _
    '   "             YOYAKU_NO," & vbCrLf & _
    '   "             COUNT(YOYAKU_NO) AS SHONIN_NINZU" & vbCrLf & _
    '   "         FROM" & vbCrLf & _
    '   "             CHECK_RIREKI_TBL" & vbCrLf & _
    '   "         WHERE" & vbCrLf & _
    '   "             CHECK_STS = '1'" & vbCrLf & _
    '   "         GROUP BY" & vbCrLf & _
    '   "             YOYAKU_NO" & vbCrLf & _
    '   "     ) t3" & vbCrLf & _
    '   " ON  t1.YOYAKU_NO = T3.YOYAKU_NO" & vbCrLf & _
    '   " LEFT JOIN" & vbCrLf & _
    '   "     (" & vbCrLf & _
    '   "         SELECT" & vbCrLf & _
    '   "             YOYAKU_NO," & vbCrLf & _
    '   "             MIN(tmp.RIYO_KIN) RIYO_KIN" & vbCrLf & _
    '   "         FROM" & vbCrLf & _
    '   "             (" & vbCrLf & _
    '   "                 SELECT" & vbCrLf & _
    '   "                     YOYAKU_NO," & vbCrLf & _
    '   "                     CASE" & vbCrLf & _
    '   "                         WHEN SEIKYU_INPUT_FLG = '0' THEN '0'" & vbCrLf & _
    '   "                         ELSE CASE" & vbCrLf & _
    '   "                             WHEN SEIKYU_IRAI_FLG = '0' THEN '1'" & vbCrLf & _
    '   "                             ELSE CASE" & vbCrLf & _
    '   "                                 WHEN NYUKIN_INPUT_FLG = '0' THEN '2'" & vbCrLf & _
    '   "                                 ELSE '3'" & vbCrLf & _
    '   "                             END" & vbCrLf & _
    '   "                         END" & vbCrLf & _
    '   "                     END RIYO_KIN" & vbCrLf & _
    '   "                 FROM" & vbCrLf & _
    '   "                     BILLPAY_TBL" & vbCrLf & _
    '   "                 WHERE" & vbCrLf & _
    '   "                     SEIKYU_NAIYO = '1'" & vbCrLf & _
    '   "                 OR  SEIKYU_NAIYO = '3'" & vbCrLf & _
    '   "             )tmp" & vbCrLf & _
    '   "        GROUP BY YOYAKU_NO" & vbCrLf & _
    '   "     ) t4a" & vbCrLf & _
    '   " ON  t1.YOYAKU_NO = t4a.YOYAKU_NO" & vbCrLf & _
    '   " LEFT JOIN" & vbCrLf & _
    '   "     (" & vbCrLf & _
    '   "         SELECT" & vbCrLf & _
    '   "             YOYAKU_NO," & vbCrLf & _
    '   "             MIN(tmp.FUTAI_SETUBI) FUTAI_SETUBI" & vbCrLf & _
    '   "         FROM" & vbCrLf & _
    '   "             (" & vbCrLf & _
    '   "                 SELECT" & vbCrLf & _
    '   "                     YOYAKU_NO," & vbCrLf & _
    '   "                     CASE" & vbCrLf & _
    '   "                         WHEN SEIKYU_INPUT_FLG = '0' THEN '0'" & vbCrLf & _
    '   "                         ELSE CASE" & vbCrLf & _
    '   "                             WHEN SEIKYU_IRAI_FLG = '0' THEN '1'" & vbCrLf & _
    '   "                             ELSE CASE" & vbCrLf & _
    '   "                                 WHEN NYUKIN_INPUT_FLG = '0' THEN '2'" & vbCrLf & _
    '   "                                 ELSE '3'" & vbCrLf & _
    '   "                             END" & vbCrLf & _
    '   "                         END" & vbCrLf & _
    '   "                     END FUTAI_SETUBI" & vbCrLf & _
    '   "                 FROM" & vbCrLf & _
    '   "                     BILLPAY_TBL" & vbCrLf & _
    '   "                 WHERE" & vbCrLf & _
    '   "                     SEIKYU_NAIYO = '2'" & vbCrLf & _
    '   "                 OR  SEIKYU_NAIYO = '3'" & vbCrLf & _
    '   "             )tmp" & vbCrLf & _
    '   "        GROUP BY YOYAKU_NO" & vbCrLf & _
    '   "     ) t4b" & vbCrLf & _
    '   " ON t1.YOYAKU_NO = t4b.YOYAKU_NO" & vbCrLf & _
    '   " LEFT JOIN BILLPAY_TBL t5" & vbCrLf & _
    '   " ON t1.YOYAKU_NO = t5.YOYAKU_NO" & vbCrLf & _
    '"WHERE" & vbCrLf & _
    '"    t1.SHISETU_KBN = '1'"
    '---2016.1.15 END ↑y.ozawa 背景色変更

    ' 2016.04.19 UPD START↓ h.hagiwara キャンセル時のステータス表示追加＆利用形状表示不具合対応（着席･スタンディング逆）
    'Private strSelectYoyakuTheatreSQL As String = _
    '"SELECT" & vbCrLf & _
    '    "DISTINCT" & vbCrLf & _
    '    "t1.YOYAKU_NO," & vbCrLf & _
    '    "CASE t1.YOYAKU_STS" & vbCrLf & _
    '    "    WHEN '1' THEN '仮'" & vbCrLf & _
    '    "    WHEN '2' THEN '決定'" & vbCrLf & _
    '    "    WHEN '3' THEN '受諾'" & vbCrLf & _
    '    "    WHEN '4' THEN '完了'" & vbCrLf & _
    '    "END YOYAKU_STS," & vbCrLf & _
    '    "t2a.YOYAKU_DT_FROM," & vbCrLf & _
    '    "t2a.YOYAKU_DT_TO," & vbCrLf & _
    '    "t2b.START_TIME," & vbCrLf & _
    '    "t2c.END_TIME," & vbCrLf & _
    '    "t1.SAIJI_NM," & vbCrLf & _
    '    "CASE t1.RIYO_TYPE" & vbCrLf & _
    '    "    WHEN '0' THEN '未定'" & vbCrLf & _
    '    "    WHEN '1' THEN 'スタンディング'" & vbCrLf & _
    '    "    WHEN '2' THEN '着席'" & vbCrLf & _
    '    "    WHEN '3' THEN '変則'" & vbCrLf & _
    '    "    WHEN '4' THEN '催事'" & vbCrLf & _
    '    "END RIYO_TYPE," & vbCrLf & _
    '    "t1.RIYO_NM," & vbCrLf & _
    '    "t1.SEKININ_NM," & vbCrLf & _
    '    "t3.SHONIN_NINZU," & vbCrLf & _
    '    "CASE t4a.RIYO_KIN" & vbCrLf & _
    '    "    WHEN '0' THEN ''" & vbCrLf & _
    '    "    WHEN '1' THEN '△'" & vbCrLf & _
    '    "    WHEN '2' THEN '〇'" & vbCrLf & _
    '    "    WHEN '3' THEN '◎'" & vbCrLf & _
    '    "END RIYO_KIN," & vbCrLf & _
    '    "CASE" & vbCrLf & _
    '    "    WHEN t1.FINPUT_STS = '1' THEN '確定'" & vbCrLf & _
    '    "END FINPUT_STS," & vbCrLf & _
    '    "CASE t4b.FUTAI_SETUBI" & vbCrLf & _
    '    "    WHEN '0' THEN ''" & vbCrLf & _
    '    "    WHEN '1' THEN '△'" & vbCrLf & _
    '    "    WHEN '2' THEN '〇'" & vbCrLf & _
    '    "    WHEN '3' THEN '◎'" & vbCrLf & _
    '    "END FUTAI_SETUBI," & vbCrLf & _
    '    "t1.YOYAKU_STS AS STS_CD" & vbCrLf & _
    '"FROM" & vbCrLf & _
    '   " YOYAKU_TBL t1" & vbCrLf & _
    '   " LEFT JOIN YDT_TBL t2" & vbCrLf & _
    '   " ON  t1.YOYAKU_NO = T2.YOYAKU_NO" & vbCrLf & _
    '   " LEFT JOIN" & vbCrLf & _
    '   "     (" & vbCrLf & _
    '   "         SELECT" & vbCrLf & _
    '   "             YOYAKU_NO," & vbCrLf & _
    '   "             MIN(YOYAKU_DT) AS YOYAKU_DT_FROM," & vbCrLf & _
    '   "             MAX(YOYAKU_DT) AS YOYAKU_DT_TO" & vbCrLf & _
    '   "         FROM" & vbCrLf & _
    '   "             YDT_TBL" & vbCrLf & _
    '   "         WHERE" & vbCrLf & _
    '   "             SHISETU_KBN = '1'" & vbCrLf & _
    '   "         GROUP BY YOYAKU_NO" & vbCrLf & _
    '   "     ) t2a" & vbCrLf & _
    '   " ON  t1.YOYAKU_NO = T2a.YOYAKU_NO" & vbCrLf & _
    '   " LEFT JOIN" & vbCrLf & _
    '   "     (   SELECT" & vbCrLf & _
    '   "             YOYAKU_NO," & vbCrLf & _
    '   "             YOYAKU_DT," & vbCrLf & _
    '   "             START_TIME" & vbCrLf & _
    '   "         FROM" & vbCrLf & _
    '   "             YDT_TBL" & vbCrLf & _
    '   "         WHERE" & vbCrLf & _
    '   "             SHISETU_KBN = '1'" & vbCrLf & _
    '   "     ) t2b" & vbCrLf & _
    '   " ON  t2a.YOYAKU_NO = T2b.YOYAKU_NO" & vbCrLf & _
    '   " AND t2a.YOYAKU_DT_FROM = T2b.YOYAKU_DT" & vbCrLf & _
    '   " LEFT JOIN" & vbCrLf & _
    '   "     (" & vbCrLf & _
    '   "         SELECT" & vbCrLf & _
    '   "             YOYAKU_NO," & vbCrLf & _
    '   "             YOYAKU_DT," & vbCrLf & _
    '   "             END_TIME" & vbCrLf & _
    '   "         FROM" & vbCrLf & _
    '   "             YDT_TBL" & vbCrLf & _
    '   "         WHERE" & vbCrLf & _
    '   "             SHISETU_KBN = '1'" & vbCrLf & _
    '   "     ) t2c" & vbCrLf & _
    '   " ON  t2a.YOYAKU_NO = T2c.YOYAKU_NO" & vbCrLf & _
    '   " AND t2a.YOYAKU_DT_TO = T2c.YOYAKU_DT" & vbCrLf & _
    '   " LEFT JOIN" & vbCrLf & _
    '   "     (" & vbCrLf & _
    '   "         SELECT" & vbCrLf & _
    '   "             YOYAKU_NO," & vbCrLf & _
    '   "             COUNT(YOYAKU_NO) AS SHONIN_NINZU" & vbCrLf & _
    '   "         FROM" & vbCrLf & _
    '   "             CHECK_RIREKI_TBL" & vbCrLf & _
    '   "         WHERE" & vbCrLf & _
    '   "             CHECK_STS = '1'" & vbCrLf & _
    '   "         GROUP BY" & vbCrLf & _
    '   "             YOYAKU_NO" & vbCrLf & _
    '   "     ) t3" & vbCrLf & _
    '   " ON  t1.YOYAKU_NO = T3.YOYAKU_NO" & vbCrLf & _
    '   " LEFT JOIN" & vbCrLf & _
    '   "     (" & vbCrLf & _
    '   "         SELECT" & vbCrLf & _
    '   "             YOYAKU_NO," & vbCrLf & _
    '   "             MIN(tmp.RIYO_KIN) RIYO_KIN" & vbCrLf & _
    '   "         FROM" & vbCrLf & _
    '   "             (" & vbCrLf & _
    '   "                 SELECT" & vbCrLf & _
    '   "                     YOYAKU_NO," & vbCrLf & _
    '   "                     CASE" & vbCrLf & _
    '   "                         WHEN SEIKYU_INPUT_FLG = '0' THEN '0'" & vbCrLf & _
    '   "                         ELSE CASE" & vbCrLf & _
    '   "                             WHEN SEIKYU_IRAI_FLG = '0' THEN '1'" & vbCrLf & _
    '   "                             ELSE CASE" & vbCrLf & _
    '   "                                 WHEN NYUKIN_INPUT_FLG = '0' THEN '2'" & vbCrLf & _
    '   "                                 ELSE '3'" & vbCrLf & _
    '   "                             END" & vbCrLf & _
    '   "                         END" & vbCrLf & _
    '   "                     END RIYO_KIN" & vbCrLf & _
    '   "                 FROM" & vbCrLf & _
    '   "                     BILLPAY_TBL" & vbCrLf & _
    '   "                 WHERE" & vbCrLf & _
    '   "                     SEIKYU_NAIYO = '1'" & vbCrLf & _
    '   "                 OR  SEIKYU_NAIYO = '3'" & vbCrLf & _
    '   "             )tmp" & vbCrLf & _
    '   "        GROUP BY YOYAKU_NO" & vbCrLf & _
    '   "     ) t4a" & vbCrLf & _
    '   " ON  t1.YOYAKU_NO = t4a.YOYAKU_NO" & vbCrLf & _
    '   " LEFT JOIN" & vbCrLf & _
    '   "     (" & vbCrLf & _
    '   "         SELECT" & vbCrLf & _
    '   "             YOYAKU_NO," & vbCrLf & _
    '   "             MIN(tmp.FUTAI_SETUBI) FUTAI_SETUBI" & vbCrLf & _
    '   "         FROM" & vbCrLf & _
    '   "             (" & vbCrLf & _
    '   "                 SELECT" & vbCrLf & _
    '   "                     YOYAKU_NO," & vbCrLf & _
    '   "                     CASE" & vbCrLf & _
    '   "                         WHEN SEIKYU_INPUT_FLG = '0' THEN '0'" & vbCrLf & _
    '   "                         ELSE CASE" & vbCrLf & _
    '   "                             WHEN SEIKYU_IRAI_FLG = '0' THEN '1'" & vbCrLf & _
    '   "                             ELSE CASE" & vbCrLf & _
    '   "                                 WHEN NYUKIN_INPUT_FLG = '0' THEN '2'" & vbCrLf & _
    '   "                                 ELSE '3'" & vbCrLf & _
    '   "                             END" & vbCrLf & _
    '   "                         END" & vbCrLf & _
    '   "                     END FUTAI_SETUBI" & vbCrLf & _
    '   "                 FROM" & vbCrLf & _
    '   "                     BILLPAY_TBL" & vbCrLf & _
    '   "                 WHERE" & vbCrLf & _
    '   "                     SEIKYU_NAIYO = '2'" & vbCrLf & _
    '   "                 OR  SEIKYU_NAIYO = '3'" & vbCrLf & _
    '   "             )tmp" & vbCrLf & _
    '   "        GROUP BY YOYAKU_NO" & vbCrLf & _
    '   "     ) t4b" & vbCrLf & _
    '   " ON t1.YOYAKU_NO = t4b.YOYAKU_NO" & vbCrLf & _
    '"WHERE" & vbCrLf & _
    '"    t1.SHISETU_KBN = '1'"
    '' 2016.04.04 UPD END↑ h.hagiwara 条件変更に伴い参照テーブル削除
    Private strSelectYoyakuTheatreSQL As String = _
    "SELECT" & vbCrLf & _
        "DISTINCT" & vbCrLf & _
        "t1.YOYAKU_NO," & vbCrLf & _
        "CASE t1.YOYAKU_STS" & vbCrLf & _
        "    WHEN '1' THEN '仮'" & vbCrLf & _
        "    WHEN '2' THEN '決定'" & vbCrLf & _
        "    WHEN '3' THEN '受諾'" & vbCrLf & _
        "    WHEN '4' THEN '完了'" & vbCrLf & _
        "    WHEN '5' THEN '仮予Ｃ'" & vbCrLf & _
        "    WHEN '6' THEN '正予Ｃ'" & vbCrLf & _
        "END YOYAKU_STS," & vbCrLf & _
        "t2a.YOYAKU_DT_FROM," & vbCrLf & _
        "t2a.YOYAKU_DT_TO," & vbCrLf & _
        "t2b.START_TIME," & vbCrLf & _
        "t2c.END_TIME," & vbCrLf & _
        "t1.SAIJI_NM," & vbCrLf & _
        "CASE t1.RIYO_TYPE" & vbCrLf & _
        "    WHEN '0' THEN '未定'" & vbCrLf & _
        "    WHEN '1' THEN '着席'" & vbCrLf & _
        "    WHEN '2' THEN 'スタンディング'" & vbCrLf & _
        "    WHEN '3' THEN '変則'" & vbCrLf & _
        "    WHEN '4' THEN '催事'" & vbCrLf & _
        "END RIYO_TYPE," & vbCrLf & _
        "t1.RIYO_NM," & vbCrLf & _
        "t1.SEKININ_NM," & vbCrLf & _
        "t3.SHONIN_NINZU," & vbCrLf & _
        "CASE t4a.RIYO_KIN" & vbCrLf & _
        "    WHEN '0' THEN ''" & vbCrLf & _
        "    WHEN '1' THEN '△'" & vbCrLf & _
        "    WHEN '2' THEN '〇'" & vbCrLf & _
        "    WHEN '3' THEN '◎'" & vbCrLf & _
        "END RIYO_KIN," & vbCrLf & _
        "CASE" & vbCrLf & _
        "    WHEN t1.FINPUT_STS = '1' THEN '確定'" & vbCrLf & _
        "END FINPUT_STS," & vbCrLf & _
        "CASE t4b.FUTAI_SETUBI" & vbCrLf & _
        "    WHEN '0' THEN ''" & vbCrLf & _
        "    WHEN '1' THEN '△'" & vbCrLf & _
        "    WHEN '2' THEN '〇'" & vbCrLf & _
        "    WHEN '3' THEN '◎'" & vbCrLf & _
        "END FUTAI_SETUBI," & vbCrLf & _
        "t1.YOYAKU_STS AS STS_CD" & vbCrLf & _
    "FROM" & vbCrLf & _
       " YOYAKU_TBL t1" & vbCrLf & _
       " LEFT JOIN YDT_TBL t2" & vbCrLf & _
       " ON  t1.YOYAKU_NO = T2.YOYAKU_NO" & vbCrLf & _
       " LEFT JOIN" & vbCrLf & _
       "     (" & vbCrLf & _
       "         SELECT" & vbCrLf & _
       "             YOYAKU_NO," & vbCrLf & _
       "             MIN(YOYAKU_DT) AS YOYAKU_DT_FROM," & vbCrLf & _
       "             MAX(YOYAKU_DT) AS YOYAKU_DT_TO" & vbCrLf & _
       "         FROM" & vbCrLf & _
       "             YDT_TBL" & vbCrLf & _
       "         WHERE" & vbCrLf & _
       "             SHISETU_KBN = '1'" & vbCrLf & _
       "         GROUP BY YOYAKU_NO" & vbCrLf & _
       "     ) t2a" & vbCrLf & _
       " ON  t1.YOYAKU_NO = T2a.YOYAKU_NO" & vbCrLf & _
       " LEFT JOIN" & vbCrLf & _
       "     (   SELECT" & vbCrLf & _
       "             YOYAKU_NO," & vbCrLf & _
       "             YOYAKU_DT," & vbCrLf & _
       "             START_TIME" & vbCrLf & _
       "         FROM" & vbCrLf & _
       "             YDT_TBL" & vbCrLf & _
       "         WHERE" & vbCrLf & _
       "             SHISETU_KBN = '1'" & vbCrLf & _
       "     ) t2b" & vbCrLf & _
       " ON  t2a.YOYAKU_NO = T2b.YOYAKU_NO" & vbCrLf & _
       " AND t2a.YOYAKU_DT_FROM = T2b.YOYAKU_DT" & vbCrLf & _
       " LEFT JOIN" & vbCrLf & _
       "     (" & vbCrLf & _
       "         SELECT" & vbCrLf & _
       "             YOYAKU_NO," & vbCrLf & _
       "             YOYAKU_DT," & vbCrLf & _
       "             END_TIME" & vbCrLf & _
       "         FROM" & vbCrLf & _
       "             YDT_TBL" & vbCrLf & _
       "         WHERE" & vbCrLf & _
       "             SHISETU_KBN = '1'" & vbCrLf & _
       "     ) t2c" & vbCrLf & _
       " ON  t2a.YOYAKU_NO = T2c.YOYAKU_NO" & vbCrLf & _
       " AND t2a.YOYAKU_DT_TO = T2c.YOYAKU_DT" & vbCrLf & _
       " LEFT JOIN" & vbCrLf & _
       "     (" & vbCrLf & _
       "         SELECT" & vbCrLf & _
       "             YOYAKU_NO," & vbCrLf & _
       "             COUNT(YOYAKU_NO) AS SHONIN_NINZU" & vbCrLf & _
       "         FROM" & vbCrLf & _
       "             CHECK_RIREKI_TBL" & vbCrLf & _
       "         WHERE" & vbCrLf & _
       "             CHECK_STS = '1'" & vbCrLf & _
       "         GROUP BY" & vbCrLf & _
       "             YOYAKU_NO" & vbCrLf & _
       "     ) t3" & vbCrLf & _
       " ON  t1.YOYAKU_NO = T3.YOYAKU_NO" & vbCrLf & _
       " LEFT JOIN" & vbCrLf & _
       "     (" & vbCrLf & _
       "         SELECT" & vbCrLf & _
       "             YOYAKU_NO," & vbCrLf & _
       "             MIN(tmp.RIYO_KIN) RIYO_KIN" & vbCrLf & _
       "         FROM" & vbCrLf & _
       "             (" & vbCrLf & _
       "                 SELECT" & vbCrLf & _
       "                     YOYAKU_NO," & vbCrLf & _
       "                     CASE" & vbCrLf & _
       "                         WHEN SEIKYU_INPUT_FLG = '0' THEN '0'" & vbCrLf & _
       "                         ELSE CASE" & vbCrLf & _
       "                             WHEN SEIKYU_IRAI_FLG = '0' THEN '1'" & vbCrLf & _
       "                             ELSE CASE" & vbCrLf & _
       "                                 WHEN NYUKIN_INPUT_FLG = '0' THEN '2'" & vbCrLf & _
       "                                 ELSE '3'" & vbCrLf & _
       "                             END" & vbCrLf & _
       "                         END" & vbCrLf & _
       "                     END RIYO_KIN" & vbCrLf & _
       "                 FROM" & vbCrLf & _
       "                     BILLPAY_TBL" & vbCrLf & _
       "                 WHERE" & vbCrLf & _
       "                     SEIKYU_NAIYO = '1'" & vbCrLf & _
       "                 OR  SEIKYU_NAIYO = '3'" & vbCrLf & _
       "             )tmp" & vbCrLf & _
       "        GROUP BY YOYAKU_NO" & vbCrLf & _
       "     ) t4a" & vbCrLf & _
       " ON  t1.YOYAKU_NO = t4a.YOYAKU_NO" & vbCrLf & _
       " LEFT JOIN" & vbCrLf & _
       "     (" & vbCrLf & _
       "         SELECT" & vbCrLf & _
       "             YOYAKU_NO," & vbCrLf & _
       "             MIN(tmp.FUTAI_SETUBI) FUTAI_SETUBI" & vbCrLf & _
       "         FROM" & vbCrLf & _
       "             (" & vbCrLf & _
       "                 SELECT" & vbCrLf & _
       "                     YOYAKU_NO," & vbCrLf & _
       "                     CASE" & vbCrLf & _
       "                         WHEN SEIKYU_INPUT_FLG = '0' THEN '0'" & vbCrLf & _
       "                         ELSE CASE" & vbCrLf & _
       "                             WHEN SEIKYU_IRAI_FLG = '0' THEN '1'" & vbCrLf & _
       "                             ELSE CASE" & vbCrLf & _
       "                                 WHEN NYUKIN_INPUT_FLG = '0' THEN '2'" & vbCrLf & _
       "                                 ELSE '3'" & vbCrLf & _
       "                             END" & vbCrLf & _
       "                         END" & vbCrLf & _
       "                     END FUTAI_SETUBI" & vbCrLf & _
       "                 FROM" & vbCrLf & _
       "                     BILLPAY_TBL" & vbCrLf & _
       "                 WHERE" & vbCrLf & _
       "                     SEIKYU_NAIYO = '2'" & vbCrLf & _
       "                 OR  SEIKYU_NAIYO = '3'" & vbCrLf & _
       "             )tmp" & vbCrLf & _
       "        GROUP BY YOYAKU_NO" & vbCrLf & _
       "     ) t4b" & vbCrLf & _
       " ON t1.YOYAKU_NO = t4b.YOYAKU_NO" & vbCrLf & _
    "WHERE" & vbCrLf & _
    "    t1.SHISETU_KBN = '1'"
    ' 2016.04.19 UPD END↑ h.hagiwara キャンセル時のステータス表示追加＆利用形状表示不具合対応（着席･スタンディング逆）

#End Region

#Region "SQL文(予約一覧（スタジオ）取得)"

    'SQL文(予約一覧（スタジオ）取得)
    ' 2016.1.15 START ↓y.ozawa
    'Private strSelectYoyakuStudioSQL As String = _
    '"SELECT" & vbCrLf & _
    '    "DISTINCT" & vbCrLf & _
    '"    t1.YOYAKU_NO," & vbCrLf & _
    '"    CASE t1.YOYAKU_STS" & vbCrLf & _
    '"        WHEN '1' THEN '仮未'" & vbCrLf & _
    '"        WHEN '2' THEN '仮'" & vbCrLf & _
    '"        WHEN '3' THEN '正式'" & vbCrLf & _
    '"        WHEN '4' THEN '完了'" & vbCrLf & _
    '"    END YOYAKU_STS," & vbCrLf & _
    '"    CASE t1.STUDIO_KBN" & vbCrLf & _
    '"        WHEN '1' THEN '201st'" & vbCrLf & _
    '"        WHEN '2' THEN '202st'" & vbCrLf & _
    '"        WHEN '3' THEN 'house lock'" & vbCrLf & _
    '"    END STUDIO," & vbCrLf & _
    '"    t2a.YOYAKU_DT_FROM," & vbCrLf & _
    '"    t2a.YOYAKU_DT_TO," & vbCrLf & _
    '"    t2b.START_TIME," & vbCrLf & _
    '"    t2c.END_TIME," & vbCrLf & _
    '"    t1.SHUTSUEN_NM," & vbCrLf & _
    '"    t1.RIYO_NM," & vbCrLf & _
    '"    t1.SEKININ_NM," & vbCrLf & _
    '"    t3.SHONIN_NINZU," & vbCrLf & _
    '"    CASE t4a.RIYO_KIN" & vbCrLf & _
    '"        WHEN '0' THEN ''" & vbCrLf & _
    '"        WHEN '1' THEN '△'" & vbCrLf & _
    '"        WHEN '2' THEN '〇'" & vbCrLf & _
    '"        WHEN '3' THEN '◎'" & vbCrLf & _
    '"    END RIYO_KIN," & vbCrLf & _
    '"    CASE" & vbCrLf & _
    '"        WHEN t1.FINPUT_STS = '1' THEN '確定'" & vbCrLf & _
    '"    END FINPUT_STS," & vbCrLf & _
    '"    CASE t4b.FUTAI_SETUBI" & vbCrLf & _
    '"        WHEN '0' THEN ''" & vbCrLf & _
    '"        WHEN '1' THEN '△'" & vbCrLf & _
    '"        WHEN '2' THEN '〇'" & vbCrLf & _
    '"        WHEN '3' THEN '◎'" & vbCrLf & _
    '"    END FUTAI_SETUBI," & vbCrLf & _
    '"    t1.YOYAKU_STS AS STS_CD" & vbCrLf & _
    '"FROM" & vbCrLf & _
    '   " YOYAKU_TBL t1" & vbCrLf & _
    '   " LEFT JOIN YDT_TBL t2" & vbCrLf & _
    '   " ON  t1.YOYAKU_NO = t2.YOYAKU_NO" & vbCrLf & _
    '   " LEFT JOIN" & vbCrLf & _
    '   "     (" & vbCrLf & _
    '   "         SELECT" & vbCrLf & _
    '   "             YOYAKU_NO," & vbCrLf & _
    '   "             MIN(YOYAKU_DT) AS YOYAKU_DT_FROM," & vbCrLf & _
    '   "             MAX(YOYAKU_DT) AS YOYAKU_DT_TO" & vbCrLf & _
    '   "         FROM" & vbCrLf & _
    '   "             YDT_TBL" & vbCrLf & _
    '   "         WHERE" & vbCrLf & _
    '   "             SHISETU_KBN = '2'" & vbCrLf & _
    '   "         GROUP BY YOYAKU_NO" & vbCrLf & _
    '   "     ) t2a" & vbCrLf & _
    '   " ON  t2.YOYAKU_NO = T2a.YOYAKU_NO" & vbCrLf & _
    '   " LEFT JOIN" & vbCrLf & _
    '   "     (   SELECT" & vbCrLf & _
    '   "             YOYAKU_NO," & vbCrLf & _
    '   "             YOYAKU_DT," & vbCrLf & _
    '   "             START_TIME" & vbCrLf & _
    '   "         FROM" & vbCrLf & _
    '   "             YDT_TBL" & vbCrLf & _
    '   "         WHERE" & vbCrLf & _
    '   "             SHISETU_KBN = '2'" & vbCrLf & _
    '   "     ) t2b" & vbCrLf & _
    '   " ON  t2a.YOYAKU_NO = t2b.YOYAKU_NO" & vbCrLf & _
    '   " AND t2a.YOYAKU_DT_FROM = t2b.YOYAKU_DT" & vbCrLf & _
    '   " LEFT JOIN" & vbCrLf & _
    '   "     (" & vbCrLf & _
    '   "         SELECT" & vbCrLf & _
    '   "             YOYAKU_NO," & vbCrLf & _
    '   "             YOYAKU_DT," & vbCrLf & _
    '   "             END_TIME" & vbCrLf & _
    '   "         FROM" & vbCrLf & _
    '   "             YDT_TBL" & vbCrLf & _
    '   "         WHERE" & vbCrLf & _
    '   "             SHISETU_KBN = '2'" & vbCrLf & _
    '    "     ) t2c" & vbCrLf & _
    '   " ON  t2a.YOYAKU_NO = t2c.YOYAKU_NO" & vbCrLf & _
    '   " AND t2a.YOYAKU_DT_TO = t2c.YOYAKU_DT" & vbCrLf & _
    '   " LEFT JOIN" & vbCrLf & _
    '   "     (" & vbCrLf & _
    '   "         SELECT" & vbCrLf & _
    '   "             YOYAKU_NO," & vbCrLf & _
    '   "             COUNT(YOYAKU_NO) AS SHONIN_NINZU" & vbCrLf & _
    '   "         FROM" & vbCrLf & _
    '   "             CHECK_RIREKI_TBL" & vbCrLf & _
    '   "         WHERE" & vbCrLf & _
    '   "             CHECK_STS = '1'" & vbCrLf & _
    '   "         GROUP BY" & vbCrLf & _
    '   "             YOYAKU_NO" & vbCrLf & _
    '   "     ) t3" & vbCrLf & _
    '   " ON  t1.YOYAKU_NO = T3.YOYAKU_NO" & vbCrLf & _
    '   " LEFT JOIN" & vbCrLf & _
    '   "     (" & vbCrLf & _
    '   "         SELECT" & vbCrLf & _
    '   "             YOYAKU_NO," & vbCrLf & _
    '   "             MIN(tmp.RIYO_KIN) RIYO_KIN" & vbCrLf & _
    '   "         FROM" & vbCrLf & _
    '   "             (" & vbCrLf & _
    '   "                 SELECT" & vbCrLf & _
    '   "                     YOYAKU_NO," & vbCrLf & _
    '   "                     CASE" & vbCrLf & _
    '   "                         WHEN SEIKYU_INPUT_FLG = '0' THEN '0'" & vbCrLf & _
    '   "                         ELSE CASE" & vbCrLf & _
    '   "                             WHEN SEIKYU_IRAI_FLG = '0' THEN '1'" & vbCrLf & _
    '   "                             ELSE CASE" & vbCrLf & _
    '   "                                 WHEN NYUKIN_INPUT_FLG = '0' THEN '2'" & vbCrLf & _
    '   "                                 ELSE '3'" & vbCrLf & _
    '   "                             END" & vbCrLf & _
    '   "                         END" & vbCrLf & _
    '   "                     END RIYO_KIN" & vbCrLf & _
    '   "                 FROM" & vbCrLf & _
    '   "                     BILLPAY_TBL" & vbCrLf & _
    '   "                 WHERE" & vbCrLf & _
    '   "                     SEIKYU_NAIYO = '1'" & vbCrLf & _
    '   "                 OR  SEIKYU_NAIYO = '3'" & vbCrLf & _
    '   "             )tmp" & vbCrLf & _
    '   "        GROUP BY YOYAKU_NO" & vbCrLf & _
    '   "     ) t4a" & vbCrLf & _
    '   " ON  t1.YOYAKU_NO = t4a.YOYAKU_NO" & vbCrLf & _
    '   " LEFT JOIN" & vbCrLf & _
    '   "     (" & vbCrLf & _
    '   "         SELECT" & vbCrLf & _
    '   "             YOYAKU_NO," & vbCrLf & _
    '   "             MIN(tmp.FUTAI_SETUBI) FUTAI_SETUBI" & vbCrLf & _
    '   "         FROM" & vbCrLf & _
    '   "             (" & vbCrLf & _
    '   "                 SELECT" & vbCrLf & _
    '   "                     YOYAKU_NO," & vbCrLf & _
    '   "                     CASE" & vbCrLf & _
    '   "                         WHEN SEIKYU_INPUT_FLG = '0' THEN '0'" & vbCrLf & _
    '   "                         ELSE CASE" & vbCrLf & _
    '   "                             WHEN SEIKYU_IRAI_FLG = '0' THEN '1'" & vbCrLf & _
    '   "                             ELSE CASE" & vbCrLf & _
    '   "                                 WHEN NYUKIN_INPUT_FLG = '0' THEN '2'" & vbCrLf & _
    '   "                                 ELSE '3'" & vbCrLf & _
    '   "                             END" & vbCrLf & _
    '   "                         END" & vbCrLf & _
    '   "                     END FUTAI_SETUBI" & vbCrLf & _
    '   "                 FROM" & vbCrLf & _
    '   "                     BILLPAY_TBL" & vbCrLf & _
    '   "                 WHERE" & vbCrLf & _
    '   "                     SEIKYU_NAIYO = '2'" & vbCrLf & _
    '   "                 OR  SEIKYU_NAIYO = '3'" & vbCrLf & _
    '   "             )tmp" & vbCrLf & _
    '   "        GROUP BY YOYAKU_NO" & vbCrLf & _
    '   "     ) t4b" & vbCrLf & _
    '   " ON t1.YOYAKU_NO = t4b.YOYAKU_NO" & vbCrLf & _
    '   " LEFT JOIN BILLPAY_TBL t5" & vbCrLf & _
    '   " ON t1.YOYAKU_NO = t5.YOYAKU_NO" & vbCrLf & _
    '"WHERE" & vbCrLf & _
    '"    t1.SHISETU_KBN = '2'"
    ' 2016.04.04 UPD START↓ h.hagiwara 条件変更に伴い参照テーブル削除
    'Private strSelectYoyakuStudioSQL As String = _
    '"SELECT" & vbCrLf & _
    '    "DISTINCT" & vbCrLf & _
    '"    t1.YOYAKU_NO," & vbCrLf & _
    '"    CASE t1.YOYAKU_STS" & vbCrLf & _
    '"        WHEN '1' THEN '仮'" & vbCrLf & _
    '"        WHEN '2' THEN '決定'" & vbCrLf & _
    '"        WHEN '3' THEN '受諾'" & vbCrLf & _
    '"        WHEN '4' THEN '完了'" & vbCrLf & _
    '"    END YOYAKU_STS," & vbCrLf & _
    '"    CASE t1.STUDIO_KBN" & vbCrLf & _
    '"        WHEN '1' THEN '201st'" & vbCrLf & _
    '"        WHEN '2' THEN '202st'" & vbCrLf & _
    '"        WHEN '3' THEN 'house lock'" & vbCrLf & _
    '"    END STUDIO," & vbCrLf & _
    '"    t2a.YOYAKU_DT_FROM," & vbCrLf & _
    '"    t2a.YOYAKU_DT_TO," & vbCrLf & _
    '"    t2b.START_TIME," & vbCrLf & _
    '"    t2c.END_TIME," & vbCrLf & _
    '"    t1.SHUTSUEN_NM," & vbCrLf & _
    '"    t1.RIYO_NM," & vbCrLf & _
    '"    t1.SEKININ_NM," & vbCrLf & _
    '"    t3.SHONIN_NINZU," & vbCrLf & _
    '"    CASE t4a.RIYO_KIN" & vbCrLf & _
    '"        WHEN '0' THEN ''" & vbCrLf & _
    '"        WHEN '1' THEN '△'" & vbCrLf & _
    '"        WHEN '2' THEN '〇'" & vbCrLf & _
    '"        WHEN '3' THEN '◎'" & vbCrLf & _
    '"    END RIYO_KIN," & vbCrLf & _
    '"    CASE" & vbCrLf & _
    '"        WHEN t1.FINPUT_STS = '1' THEN '確定'" & vbCrLf & _
    '"    END FINPUT_STS," & vbCrLf & _
    '"    CASE t4b.FUTAI_SETUBI" & vbCrLf & _
    '"        WHEN '0' THEN ''" & vbCrLf & _
    '"        WHEN '1' THEN '△'" & vbCrLf & _
    '"        WHEN '2' THEN '〇'" & vbCrLf & _
    '"        WHEN '3' THEN '◎'" & vbCrLf & _
    '"    END FUTAI_SETUBI," & vbCrLf & _
    '"    t1.YOYAKU_STS AS STS_CD" & vbCrLf & _
    '"FROM" & vbCrLf & _
    '   " YOYAKU_TBL t1" & vbCrLf & _
    '   " LEFT JOIN YDT_TBL t2" & vbCrLf & _
    '   " ON  t1.YOYAKU_NO = t2.YOYAKU_NO" & vbCrLf & _
    '   " LEFT JOIN" & vbCrLf & _
    '   "     (" & vbCrLf & _
    '   "         SELECT" & vbCrLf & _
    '   "             YOYAKU_NO," & vbCrLf & _
    '   "             MIN(YOYAKU_DT) AS YOYAKU_DT_FROM," & vbCrLf & _
    '   "             MAX(YOYAKU_DT) AS YOYAKU_DT_TO" & vbCrLf & _
    '   "         FROM" & vbCrLf & _
    '   "             YDT_TBL" & vbCrLf & _
    '   "         WHERE" & vbCrLf & _
    '   "             SHISETU_KBN = '2'" & vbCrLf & _
    '   "         GROUP BY YOYAKU_NO" & vbCrLf & _
    '   "     ) t2a" & vbCrLf & _
    '   " ON  t2.YOYAKU_NO = T2a.YOYAKU_NO" & vbCrLf & _
    '   " LEFT JOIN" & vbCrLf & _
    '   "     (   SELECT" & vbCrLf & _
    '   "             YOYAKU_NO," & vbCrLf & _
    '   "             YOYAKU_DT," & vbCrLf & _
    '   "             START_TIME" & vbCrLf & _
    '   "         FROM" & vbCrLf & _
    '   "             YDT_TBL" & vbCrLf & _
    '   "         WHERE" & vbCrLf & _
    '   "             SHISETU_KBN = '2'" & vbCrLf & _
    '   "     ) t2b" & vbCrLf & _
    '   " ON  t2a.YOYAKU_NO = t2b.YOYAKU_NO" & vbCrLf & _
    '   " AND t2a.YOYAKU_DT_FROM = t2b.YOYAKU_DT" & vbCrLf & _
    '   " LEFT JOIN" & vbCrLf & _
    '   "     (" & vbCrLf & _
    '   "         SELECT" & vbCrLf & _
    '   "             YOYAKU_NO," & vbCrLf & _
    '   "             YOYAKU_DT," & vbCrLf & _
    '   "             END_TIME" & vbCrLf & _
    '   "         FROM" & vbCrLf & _
    '   "             YDT_TBL" & vbCrLf & _
    '   "         WHERE" & vbCrLf & _
    '   "             SHISETU_KBN = '2'" & vbCrLf & _
    '    "     ) t2c" & vbCrLf & _
    '   " ON  t2a.YOYAKU_NO = t2c.YOYAKU_NO" & vbCrLf & _
    '   " AND t2a.YOYAKU_DT_TO = t2c.YOYAKU_DT" & vbCrLf & _
    '   " LEFT JOIN" & vbCrLf & _
    '   "     (" & vbCrLf & _
    '   "         SELECT" & vbCrLf & _
    '   "             YOYAKU_NO," & vbCrLf & _
    '   "             COUNT(YOYAKU_NO) AS SHONIN_NINZU" & vbCrLf & _
    '   "         FROM" & vbCrLf & _
    '   "             CHECK_RIREKI_TBL" & vbCrLf & _
    '   "         WHERE" & vbCrLf & _
    '   "             CHECK_STS = '1'" & vbCrLf & _
    '   "         GROUP BY" & vbCrLf & _
    '   "             YOYAKU_NO" & vbCrLf & _
    '   "     ) t3" & vbCrLf & _
    '   " ON  t1.YOYAKU_NO = T3.YOYAKU_NO" & vbCrLf & _
    '   " LEFT JOIN" & vbCrLf & _
    '   "     (" & vbCrLf & _
    '   "         SELECT" & vbCrLf & _
    '   "             YOYAKU_NO," & vbCrLf & _
    '   "             MIN(tmp.RIYO_KIN) RIYO_KIN" & vbCrLf & _
    '   "         FROM" & vbCrLf & _
    '   "             (" & vbCrLf & _
    '   "                 SELECT" & vbCrLf & _
    '   "                     YOYAKU_NO," & vbCrLf & _
    '   "                     CASE" & vbCrLf & _
    '   "                         WHEN SEIKYU_INPUT_FLG = '0' THEN '0'" & vbCrLf & _
    '   "                         ELSE CASE" & vbCrLf & _
    '   "                             WHEN SEIKYU_IRAI_FLG = '0' THEN '1'" & vbCrLf & _
    '   "                             ELSE CASE" & vbCrLf & _
    '   "                                 WHEN NYUKIN_INPUT_FLG = '0' THEN '2'" & vbCrLf & _
    '   "                                 ELSE '3'" & vbCrLf & _
    '   "                             END" & vbCrLf & _
    '   "                         END" & vbCrLf & _
    '   "                     END RIYO_KIN" & vbCrLf & _
    '   "                 FROM" & vbCrLf & _
    '   "                     BILLPAY_TBL" & vbCrLf & _
    '   "                 WHERE" & vbCrLf & _
    '   "                     SEIKYU_NAIYO = '1'" & vbCrLf & _
    '   "                 OR  SEIKYU_NAIYO = '3'" & vbCrLf & _
    '   "             )tmp" & vbCrLf & _
    '   "        GROUP BY YOYAKU_NO" & vbCrLf & _
    '   "     ) t4a" & vbCrLf & _
    '   " ON  t1.YOYAKU_NO = t4a.YOYAKU_NO" & vbCrLf & _
    '   " LEFT JOIN" & vbCrLf & _
    '   "     (" & vbCrLf & _
    '   "         SELECT" & vbCrLf & _
    '   "             YOYAKU_NO," & vbCrLf & _
    '   "             MIN(tmp.FUTAI_SETUBI) FUTAI_SETUBI" & vbCrLf & _
    '   "         FROM" & vbCrLf & _
    '   "             (" & vbCrLf & _
    '   "                 SELECT" & vbCrLf & _
    '   "                     YOYAKU_NO," & vbCrLf & _
    '   "                     CASE" & vbCrLf & _
    '   "                         WHEN SEIKYU_INPUT_FLG = '0' THEN '0'" & vbCrLf & _
    '   "                         ELSE CASE" & vbCrLf & _
    '   "                             WHEN SEIKYU_IRAI_FLG = '0' THEN '1'" & vbCrLf & _
    '   "                             ELSE CASE" & vbCrLf & _
    '   "                                 WHEN NYUKIN_INPUT_FLG = '0' THEN '2'" & vbCrLf & _
    '   "                                 ELSE '3'" & vbCrLf & _
    '   "                             END" & vbCrLf & _
    '   "                         END" & vbCrLf & _
    '   "                     END FUTAI_SETUBI" & vbCrLf & _
    '   "                 FROM" & vbCrLf & _
    '   "                     BILLPAY_TBL" & vbCrLf & _
    '   "                 WHERE" & vbCrLf & _
    '   "                     SEIKYU_NAIYO = '2'" & vbCrLf & _
    '   "                 OR  SEIKYU_NAIYO = '3'" & vbCrLf & _
    '   "             )tmp" & vbCrLf & _
    '   "        GROUP BY YOYAKU_NO" & vbCrLf & _
    '   "     ) t4b" & vbCrLf & _
    '   " ON t1.YOYAKU_NO = t4b.YOYAKU_NO" & vbCrLf & _
    '   " LEFT JOIN BILLPAY_TBL t5" & vbCrLf & _
    '   " ON t1.YOYAKU_NO = t5.YOYAKU_NO" & vbCrLf & _
    '"WHERE" & vbCrLf & _
    '"    t1.SHISETU_KBN = '2'"
    '---2016.1.15 END ↑y.ozawa 背景色変更

    ' 2016.04.19 UPD START↓ h.hagiwara キャンセル時のステータス表示追加
    'Private strSelectYoyakuStudioSQL As String = _
    '"SELECT" & vbCrLf & _
    '    "DISTINCT" & vbCrLf & _
    '"    t1.YOYAKU_NO," & vbCrLf & _
    '"    CASE t1.YOYAKU_STS" & vbCrLf & _
    '"        WHEN '1' THEN '仮'" & vbCrLf & _
    '"        WHEN '2' THEN '決定'" & vbCrLf & _
    '"        WHEN '3' THEN '受諾'" & vbCrLf & _
    '"        WHEN '4' THEN '完了'" & vbCrLf & _
    '"    END YOYAKU_STS," & vbCrLf & _
    '"    CASE t1.STUDIO_KBN" & vbCrLf & _
    '"        WHEN '1' THEN '201st'" & vbCrLf & _
    '"        WHEN '2' THEN '202st'" & vbCrLf & _
    '"        WHEN '3' THEN 'house lock'" & vbCrLf & _
    '"    END STUDIO," & vbCrLf & _
    '"    t2a.YOYAKU_DT_FROM," & vbCrLf & _
    '"    t2a.YOYAKU_DT_TO," & vbCrLf & _
    '"    t2b.START_TIME," & vbCrLf & _
    '"    t2c.END_TIME," & vbCrLf & _
    '"    t1.SHUTSUEN_NM," & vbCrLf & _
    '"    t1.RIYO_NM," & vbCrLf & _
    '"    t1.SEKININ_NM," & vbCrLf & _
    '"    t3.SHONIN_NINZU," & vbCrLf & _
    '"    CASE t4a.RIYO_KIN" & vbCrLf & _
    '"        WHEN '0' THEN ''" & vbCrLf & _
    '"        WHEN '1' THEN '△'" & vbCrLf & _
    '"        WHEN '2' THEN '〇'" & vbCrLf & _
    '"        WHEN '3' THEN '◎'" & vbCrLf & _
    '"    END RIYO_KIN," & vbCrLf & _
    '"    CASE" & vbCrLf & _
    '"        WHEN t1.FINPUT_STS = '1' THEN '確定'" & vbCrLf & _
    '"    END FINPUT_STS," & vbCrLf & _
    '"    CASE t4b.FUTAI_SETUBI" & vbCrLf & _
    '"        WHEN '0' THEN ''" & vbCrLf & _
    '"        WHEN '1' THEN '△'" & vbCrLf & _
    '"        WHEN '2' THEN '〇'" & vbCrLf & _
    '"        WHEN '3' THEN '◎'" & vbCrLf & _
    '"    END FUTAI_SETUBI," & vbCrLf & _
    '"    t1.YOYAKU_STS AS STS_CD" & vbCrLf & _
    '"FROM" & vbCrLf & _
    '   " YOYAKU_TBL t1" & vbCrLf & _
    '   " LEFT JOIN YDT_TBL t2" & vbCrLf & _
    '   " ON  t1.YOYAKU_NO = t2.YOYAKU_NO" & vbCrLf & _
    '   " LEFT JOIN" & vbCrLf & _
    '   "     (" & vbCrLf & _
    '   "         SELECT" & vbCrLf & _
    '   "             YOYAKU_NO," & vbCrLf & _
    '   "             MIN(YOYAKU_DT) AS YOYAKU_DT_FROM," & vbCrLf & _
    '   "             MAX(YOYAKU_DT) AS YOYAKU_DT_TO" & vbCrLf & _
    '   "         FROM" & vbCrLf & _
    '   "             YDT_TBL" & vbCrLf & _
    '   "         WHERE" & vbCrLf & _
    '   "             SHISETU_KBN = '2'" & vbCrLf & _
    '   "         GROUP BY YOYAKU_NO" & vbCrLf & _
    '   "     ) t2a" & vbCrLf & _
    '   " ON  t2.YOYAKU_NO = T2a.YOYAKU_NO" & vbCrLf & _
    '   " LEFT JOIN" & vbCrLf & _
    '   "     (   SELECT" & vbCrLf & _
    '   "             YOYAKU_NO," & vbCrLf & _
    '   "             YOYAKU_DT," & vbCrLf & _
    '   "             START_TIME" & vbCrLf & _
    '   "         FROM" & vbCrLf & _
    '   "             YDT_TBL" & vbCrLf & _
    '   "         WHERE" & vbCrLf & _
    '   "             SHISETU_KBN = '2'" & vbCrLf & _
    '   "     ) t2b" & vbCrLf & _
    '   " ON  t2a.YOYAKU_NO = t2b.YOYAKU_NO" & vbCrLf & _
    '   " AND t2a.YOYAKU_DT_FROM = t2b.YOYAKU_DT" & vbCrLf & _
    '   " LEFT JOIN" & vbCrLf & _
    '   "     (" & vbCrLf & _
    '   "         SELECT" & vbCrLf & _
    '   "             YOYAKU_NO," & vbCrLf & _
    '   "             YOYAKU_DT," & vbCrLf & _
    '   "             END_TIME" & vbCrLf & _
    '   "         FROM" & vbCrLf & _
    '   "             YDT_TBL" & vbCrLf & _
    '   "         WHERE" & vbCrLf & _
    '   "             SHISETU_KBN = '2'" & vbCrLf & _
    '    "     ) t2c" & vbCrLf & _
    '   " ON  t2a.YOYAKU_NO = t2c.YOYAKU_NO" & vbCrLf & _
    '   " AND t2a.YOYAKU_DT_TO = t2c.YOYAKU_DT" & vbCrLf & _
    '   " LEFT JOIN" & vbCrLf & _
    '   "     (" & vbCrLf & _
    '   "         SELECT" & vbCrLf & _
    '   "             YOYAKU_NO," & vbCrLf & _
    '   "             COUNT(YOYAKU_NO) AS SHONIN_NINZU" & vbCrLf & _
    '   "         FROM" & vbCrLf & _
    '   "             CHECK_RIREKI_TBL" & vbCrLf & _
    '   "         WHERE" & vbCrLf & _
    '   "             CHECK_STS = '1'" & vbCrLf & _
    '   "         GROUP BY" & vbCrLf & _
    '   "             YOYAKU_NO" & vbCrLf & _
    '   "     ) t3" & vbCrLf & _
    '   " ON  t1.YOYAKU_NO = T3.YOYAKU_NO" & vbCrLf & _
    '   " LEFT JOIN" & vbCrLf & _
    '   "     (" & vbCrLf & _
    '   "         SELECT" & vbCrLf & _
    '   "             YOYAKU_NO," & vbCrLf & _
    '   "             MIN(tmp.RIYO_KIN) RIYO_KIN" & vbCrLf & _
    '   "         FROM" & vbCrLf & _
    '   "             (" & vbCrLf & _
    '   "                 SELECT" & vbCrLf & _
    '   "                     YOYAKU_NO," & vbCrLf & _
    '   "                     CASE" & vbCrLf & _
    '   "                         WHEN SEIKYU_INPUT_FLG = '0' THEN '0'" & vbCrLf & _
    '   "                         ELSE CASE" & vbCrLf & _
    '   "                             WHEN SEIKYU_IRAI_FLG = '0' THEN '1'" & vbCrLf & _
    '   "                             ELSE CASE" & vbCrLf & _
    '   "                                 WHEN NYUKIN_INPUT_FLG = '0' THEN '2'" & vbCrLf & _
    '   "                                 ELSE '3'" & vbCrLf & _
    '   "                             END" & vbCrLf & _
    '   "                         END" & vbCrLf & _
    '   "                     END RIYO_KIN" & vbCrLf & _
    '   "                 FROM" & vbCrLf & _
    '   "                     BILLPAY_TBL" & vbCrLf & _
    '   "                 WHERE" & vbCrLf & _
    '   "                     SEIKYU_NAIYO = '1'" & vbCrLf & _
    '   "                 OR  SEIKYU_NAIYO = '3'" & vbCrLf & _
    '   "             )tmp" & vbCrLf & _
    '   "        GROUP BY YOYAKU_NO" & vbCrLf & _
    '   "     ) t4a" & vbCrLf & _
    '   " ON  t1.YOYAKU_NO = t4a.YOYAKU_NO" & vbCrLf & _
    '   " LEFT JOIN" & vbCrLf & _
    '   "     (" & vbCrLf & _
    '   "         SELECT" & vbCrLf & _
    '   "             YOYAKU_NO," & vbCrLf & _
    '   "             MIN(tmp.FUTAI_SETUBI) FUTAI_SETUBI" & vbCrLf & _
    '   "         FROM" & vbCrLf & _
    '   "             (" & vbCrLf & _
    '   "                 SELECT" & vbCrLf & _
    '   "                     YOYAKU_NO," & vbCrLf & _
    '   "                     CASE" & vbCrLf & _
    '   "                         WHEN SEIKYU_INPUT_FLG = '0' THEN '0'" & vbCrLf & _
    '   "                         ELSE CASE" & vbCrLf & _
    '   "                             WHEN SEIKYU_IRAI_FLG = '0' THEN '1'" & vbCrLf & _
    '   "                             ELSE CASE" & vbCrLf & _
    '   "                                 WHEN NYUKIN_INPUT_FLG = '0' THEN '2'" & vbCrLf & _
    '   "                                 ELSE '3'" & vbCrLf & _
    '   "                             END" & vbCrLf & _
    '   "                         END" & vbCrLf & _
    '   "                     END FUTAI_SETUBI" & vbCrLf & _
    '   "                 FROM" & vbCrLf & _
    '   "                     BILLPAY_TBL" & vbCrLf & _
    '   "                 WHERE" & vbCrLf & _
    '   "                     SEIKYU_NAIYO = '2'" & vbCrLf & _
    '   "                 OR  SEIKYU_NAIYO = '3'" & vbCrLf & _
    '   "             )tmp" & vbCrLf & _
    '   "        GROUP BY YOYAKU_NO" & vbCrLf & _
    '   "     ) t4b" & vbCrLf & _
    '   " ON t1.YOYAKU_NO = t4b.YOYAKU_NO" & vbCrLf & _
    '"WHERE" & vbCrLf & _
    '"    t1.SHISETU_KBN = '2'"
    '' 2016.04.04 UPD END↑ h.hagiwara 条件変更に伴い参照テーブル削除
    Private strSelectYoyakuStudioSQL As String = _
    "SELECT" & vbCrLf & _
        "DISTINCT" & vbCrLf & _
    "    t1.YOYAKU_NO," & vbCrLf & _
    "    CASE t1.YOYAKU_STS" & vbCrLf & _
    "        WHEN '1' THEN '仮'" & vbCrLf & _
    "        WHEN '2' THEN '決定'" & vbCrLf & _
    "        WHEN '3' THEN '受諾'" & vbCrLf & _
    "        WHEN '4' THEN '完了'" & vbCrLf & _
    "        WHEN '5' THEN '仮予Ｃ'" & vbCrLf & _
    "        WHEN '6' THEN '正予Ｃ'" & vbCrLf & _
    "    END YOYAKU_STS," & vbCrLf & _
    "    CASE t1.STUDIO_KBN" & vbCrLf & _
    "        WHEN '1' THEN '201st'" & vbCrLf & _
    "        WHEN '2' THEN '202st'" & vbCrLf & _
    "        WHEN '3' THEN 'house lock'" & vbCrLf & _
    "    END STUDIO," & vbCrLf & _
    "    t2a.YOYAKU_DT_FROM," & vbCrLf & _
    "    t2a.YOYAKU_DT_TO," & vbCrLf & _
    "    t2b.START_TIME," & vbCrLf & _
    "    t2c.END_TIME," & vbCrLf & _
    "    t1.SHUTSUEN_NM," & vbCrLf & _
    "    t1.RIYO_NM," & vbCrLf & _
    "    t1.SEKININ_NM," & vbCrLf & _
    "    t3.SHONIN_NINZU," & vbCrLf & _
    "    CASE t4a.RIYO_KIN" & vbCrLf & _
    "        WHEN '0' THEN ''" & vbCrLf & _
    "        WHEN '1' THEN '△'" & vbCrLf & _
    "        WHEN '2' THEN '〇'" & vbCrLf & _
    "        WHEN '3' THEN '◎'" & vbCrLf & _
    "    END RIYO_KIN," & vbCrLf & _
    "    CASE" & vbCrLf & _
    "        WHEN t1.FINPUT_STS = '1' THEN '確定'" & vbCrLf & _
    "    END FINPUT_STS," & vbCrLf & _
    "    CASE t4b.FUTAI_SETUBI" & vbCrLf & _
    "        WHEN '0' THEN ''" & vbCrLf & _
    "        WHEN '1' THEN '△'" & vbCrLf & _
    "        WHEN '2' THEN '〇'" & vbCrLf & _
    "        WHEN '3' THEN '◎'" & vbCrLf & _
    "    END FUTAI_SETUBI," & vbCrLf & _
    "    t1.YOYAKU_STS AS STS_CD" & vbCrLf & _
    "FROM" & vbCrLf & _
       " YOYAKU_TBL t1" & vbCrLf & _
       " LEFT JOIN YDT_TBL t2" & vbCrLf & _
       " ON  t1.YOYAKU_NO = t2.YOYAKU_NO" & vbCrLf & _
       " LEFT JOIN" & vbCrLf & _
       "     (" & vbCrLf & _
       "         SELECT" & vbCrLf & _
       "             YOYAKU_NO," & vbCrLf & _
       "             MIN(YOYAKU_DT) AS YOYAKU_DT_FROM," & vbCrLf & _
       "             MAX(YOYAKU_DT) AS YOYAKU_DT_TO" & vbCrLf & _
       "         FROM" & vbCrLf & _
       "             YDT_TBL" & vbCrLf & _
       "         WHERE" & vbCrLf & _
       "             SHISETU_KBN = '2'" & vbCrLf & _
       "         GROUP BY YOYAKU_NO" & vbCrLf & _
       "     ) t2a" & vbCrLf & _
       " ON  t2.YOYAKU_NO = T2a.YOYAKU_NO" & vbCrLf & _
       " LEFT JOIN" & vbCrLf & _
       "     (   SELECT" & vbCrLf & _
       "             YOYAKU_NO," & vbCrLf & _
       "             YOYAKU_DT," & vbCrLf & _
       "             START_TIME" & vbCrLf & _
       "         FROM" & vbCrLf & _
       "             YDT_TBL" & vbCrLf & _
       "         WHERE" & vbCrLf & _
       "             SHISETU_KBN = '2'" & vbCrLf & _
       "     ) t2b" & vbCrLf & _
       " ON  t2a.YOYAKU_NO = t2b.YOYAKU_NO" & vbCrLf & _
       " AND t2a.YOYAKU_DT_FROM = t2b.YOYAKU_DT" & vbCrLf & _
       " LEFT JOIN" & vbCrLf & _
       "     (" & vbCrLf & _
       "         SELECT" & vbCrLf & _
       "             YOYAKU_NO," & vbCrLf & _
       "             YOYAKU_DT," & vbCrLf & _
       "             END_TIME" & vbCrLf & _
       "         FROM" & vbCrLf & _
       "             YDT_TBL" & vbCrLf & _
       "         WHERE" & vbCrLf & _
       "             SHISETU_KBN = '2'" & vbCrLf & _
        "     ) t2c" & vbCrLf & _
       " ON  t2a.YOYAKU_NO = t2c.YOYAKU_NO" & vbCrLf & _
       " AND t2a.YOYAKU_DT_TO = t2c.YOYAKU_DT" & vbCrLf & _
       " LEFT JOIN" & vbCrLf & _
       "     (" & vbCrLf & _
       "         SELECT" & vbCrLf & _
       "             YOYAKU_NO," & vbCrLf & _
       "             COUNT(YOYAKU_NO) AS SHONIN_NINZU" & vbCrLf & _
       "         FROM" & vbCrLf & _
       "             CHECK_RIREKI_TBL" & vbCrLf & _
       "         WHERE" & vbCrLf & _
       "             CHECK_STS = '1'" & vbCrLf & _
       "         GROUP BY" & vbCrLf & _
       "             YOYAKU_NO" & vbCrLf & _
       "     ) t3" & vbCrLf & _
       " ON  t1.YOYAKU_NO = T3.YOYAKU_NO" & vbCrLf & _
       " LEFT JOIN" & vbCrLf & _
       "     (" & vbCrLf & _
       "         SELECT" & vbCrLf & _
       "             YOYAKU_NO," & vbCrLf & _
       "             MIN(tmp.RIYO_KIN) RIYO_KIN" & vbCrLf & _
       "         FROM" & vbCrLf & _
       "             (" & vbCrLf & _
       "                 SELECT" & vbCrLf & _
       "                     YOYAKU_NO," & vbCrLf & _
       "                     CASE" & vbCrLf & _
       "                         WHEN SEIKYU_INPUT_FLG = '0' THEN '0'" & vbCrLf & _
       "                         ELSE CASE" & vbCrLf & _
       "                             WHEN SEIKYU_IRAI_FLG = '0' THEN '1'" & vbCrLf & _
       "                             ELSE CASE" & vbCrLf & _
       "                                 WHEN NYUKIN_INPUT_FLG = '0' THEN '2'" & vbCrLf & _
       "                                 ELSE '3'" & vbCrLf & _
       "                             END" & vbCrLf & _
       "                         END" & vbCrLf & _
       "                     END RIYO_KIN" & vbCrLf & _
       "                 FROM" & vbCrLf & _
       "                     BILLPAY_TBL" & vbCrLf & _
       "                 WHERE" & vbCrLf & _
       "                     SEIKYU_NAIYO = '1'" & vbCrLf & _
       "                 OR  SEIKYU_NAIYO = '3'" & vbCrLf & _
       "             )tmp" & vbCrLf & _
       "        GROUP BY YOYAKU_NO" & vbCrLf & _
       "     ) t4a" & vbCrLf & _
       " ON  t1.YOYAKU_NO = t4a.YOYAKU_NO" & vbCrLf & _
       " LEFT JOIN" & vbCrLf & _
       "     (" & vbCrLf & _
       "         SELECT" & vbCrLf & _
       "             YOYAKU_NO," & vbCrLf & _
       "             MIN(tmp.FUTAI_SETUBI) FUTAI_SETUBI" & vbCrLf & _
       "         FROM" & vbCrLf & _
       "             (" & vbCrLf & _
       "                 SELECT" & vbCrLf & _
       "                     YOYAKU_NO," & vbCrLf & _
       "                     CASE" & vbCrLf & _
       "                         WHEN SEIKYU_INPUT_FLG = '0' THEN '0'" & vbCrLf & _
       "                         ELSE CASE" & vbCrLf & _
       "                             WHEN SEIKYU_IRAI_FLG = '0' THEN '1'" & vbCrLf & _
       "                             ELSE CASE" & vbCrLf & _
       "                                 WHEN NYUKIN_INPUT_FLG = '0' THEN '2'" & vbCrLf & _
       "                                 ELSE '3'" & vbCrLf & _
       "                             END" & vbCrLf & _
       "                         END" & vbCrLf & _
       "                     END FUTAI_SETUBI" & vbCrLf & _
       "                 FROM" & vbCrLf & _
       "                     BILLPAY_TBL" & vbCrLf & _
       "                 WHERE" & vbCrLf & _
       "                     SEIKYU_NAIYO = '2'" & vbCrLf & _
       "                 OR  SEIKYU_NAIYO = '3'" & vbCrLf & _
       "             )tmp" & vbCrLf & _
       "        GROUP BY YOYAKU_NO" & vbCrLf & _
       "     ) t4b" & vbCrLf & _
       " ON t1.YOYAKU_NO = t4b.YOYAKU_NO" & vbCrLf & _
    "WHERE" & vbCrLf & _
    "    t1.SHISETU_KBN = '2'"
    ' 2016.04.19 UPD END↑ h.hagiwara キャンセル時のステータス表示追加

#End Region

    '*****条件追加・変更、アダプタ・パラメータ網羅によるSQL文*****
#Region "予約一覧取得"

    ''' <summary>
    ''' 予約一覧取得
    ''' </summary>
    ''' <param name="Adapter">アダプタ</param>
    ''' <param name="Cn">コネクション</param>
    ''' <param name="dataEXTZ0101">データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>予約一覧を取得するSQLの作成
    ''' <para>作成情報：2015/09/09 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function setSelectYoyaku(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByRef Cn As NpgsqlConnection, _
                                       ByVal dataEXTZ0101 As DataEXTZ0101) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = String.Empty

        Try

            With dataEXTZ0101
                'SQL文(SELECT)
                'データアダプタに、SQLのSELECT文を設定
                If .PropStrShisetsuKbn = SHISETU_KBN_THEATER Then
                    'シアターの場合
                    strSQL = strSelectYoyakuTheatreSQL
                Else
                    'スタジオの場合
                    strSQL = strSelectYoyakuStudioSQL
                End If

                '利用日(From）が入力されている場合
                If .PropStrRiyoDtFrom <> String.Empty Then
                    strSQL &= "AND TO_DATE(t2.YOYAKU_DT, 'YYYY/MM/DD') >= TO_DATE('" & .PropStrRiyoDtFrom & "', 'YYYY/MM/DD') " & vbCrLf
                End If

                '利用日(To）が入力されている場合
                If .PropStrRiyoDtTo <> String.Empty Then
                    strSQL &= "AND TO_DATE(t2.YOYAKU_DT, 'YYYY/MM/DD') <= TO_DATE('" & .PropStrRiyoDtTo & "', 'YYYY/MM/DD') " & vbCrLf
                End If

                '利用者名が入力されている場合
                If .PropStrRiyoNm <> String.Empty Then
                    strSQL &= "AND t1.RIYO_NM LIKE '%' || '" & .PropStrRiyoNm & "' || '%'" & vbCrLf
                End If

                '利用者名カナが入力されている場合
                If .PropStrRiyokana <> String.Empty Then
                    strSQL &= "AND t1.RIYO_KANA LIKE '%' || '" & .PropStrRiyokana & "' || '%'" & vbCrLf
                End If

                '催事名が入力されている場合
                If .PropStrSaijiNm <> String.Empty Then
                    If .PropStrShisetsuKbn = SHISETU_KBN_THEATER Then
                        'シアターの場合
                        strSQL &= "AND t1.SAIJI_NM LIKE '%' || '" & .PropStrSaijiNm & "' || '%'" & vbCrLf
                    Else
                        'スタジオの場合
                        strSQL &= "AND t1.SHUTSUEN_NM LIKE '%' || '" & .PropStrSaijiNm & "' || '%'" & vbCrLf
                    End If
                End If

                '予約NOが入力されている場合
                If .PropStrYoyakuNo <> String.Empty Then
                    strSQL &= "AND t1.YOYAKU_NO = '" & .PropStrYoyakuNo & "' " & vbCrLf
                End If

                '未完了予約に関する出力判定
                If .PropBlnMikanryo Then
                    '未完了予約のみ出力
                    'strSQL &= " AND (t5.NYUKIN_INPUT_FLG = '0' OR t5.NYUKIN_INPUT_FLG IS NULL)"                         ' 2016.04.04 UPD h.hagiwara
                    strSQL &= " AND (t1.YOYAKU_STS <> '4')"                                                              ' 2016.04.04 UPD h.hagiwara
                Else
                    '全予約出力
                End If

                ' 2015.12.04 ADD START↓ h.hagiwara ソート順設定
                strSQL &= " ORDER BY YOYAKU_DT_FROM DESC ,  YOYAKU_NO "
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
