Imports Common
Imports CommonEXT
Imports Npgsql

''' <summary>
''' 日別売上一覧で使用する情報取得SQL作成
''' </summary>
''' <remarks>日別売上一覧で使用する情報取得のSQLを作成する
''' <para>作成情報：2015/09/16 h.endo
''' <p>改訂情報:2015/10/20 m.hayabuchi</p>
''' </para></remarks>
Public Class SqlEXTZ0103

    '*****基本SQL文*****
#Region "SQL文(日別売上一覧(シアター)取得)"

    'SQL文(日別売上一覧(シアター)取得)
    ' 2016.02.15 UPD START↓ h.hagiwara レジごとに金額集計を行い連結するよう対応
    '    Private strSelectDayUriageTheatreSQL As String = _
    '"SELECT" & vbCrLf & _
    '"    t1.YOYAKU_NO," & vbCrLf & _
    '"    t2.YOYAKU_DT," & vbCrLf & _
    '"    t1.SAIJI_NM," & vbCrLf & _
    '"    t1.RIYO_NM," & vbCrLf & _
    '"    CASE t1.KASHI_KIND" & vbCrLf & _
    '"        WHEN '0' THEN '未定'" & vbCrLf & _
    '"        WHEN '1' THEN '一般貸出'" & vbCrLf & _
    '"        WHEN '2' THEN '社内利用'" & vbCrLf & _
    '"        WHEN '3' THEN '特例'" & vbCrLf & _
    '"    END KASHI_KIND," & vbCrLf & _
    '"    CASE t1.RIYO_TYPE" & vbCrLf & _
    '"        WHEN '0' THEN '未定'" & vbCrLf & _
    '"        WHEN '1' THEN 'スタンディング'" & vbCrLf & _
    '"        WHEN '2' THEN '着席'" & vbCrLf & _
    '"        WHEN '3' THEN '変則'" & vbCrLf & _
    '"        WHEN '4' THEN '催事'" & vbCrLf & _
    '"    END RIYO_TYPE," & vbCrLf & _
    '"    CASE t1.SAIJI_BUNRUI" & vbCrLf & _
    '"        WHEN '0' THEN '未定'" & vbCrLf & _
    '"        WHEN '1' THEN '音楽'" & vbCrLf & _
    '"        WHEN '2' THEN '演劇'" & vbCrLf & _
    '"        WHEN '3' THEN '演芸'" & vbCrLf & _
    '"        WHEN '4' THEN 'ビジネス'" & vbCrLf & _
    '"        WHEN '5' THEN '試写会・映画'" & vbCrLf & _
    '"        WHEN '6' THEN 'その他'" & vbCrLf & _
    '"    END SAIJI_BUNRUI," & vbCrLf & _
    '"    COALESCE(t2.RIYO_KIN,'0') AS RIYO_KIN," & vbCrLf & _
    '"    COALESCE(t3.FUTAI_KIN,'0') AS FUTAI_KIN," & vbCrLf & _
    '"    COALESCE(t3.FUTAI_CHOSEI,'0') AS FUTAI_CHOSEI," & vbCrLf & _
    '"    COALESCE(t4a.DEPOSIT_AMOUNT,'0') AS DRINK_GENKIN," & vbCrLf & _
    '"    COALESCE(t4b.DEPOSIT_AMOUNT,'0') AS ONE_DRINK," & vbCrLf & _
    '"    COALESCE(t4c.DEPOSIT_AMOUNT,'0') AS COIN_LOCKER," & vbCrLf & _
    '"    COALESCE(t4d.DEPOSIT_AMOUNT,'0') AS ZATSU_SHUNYU," & vbCrLf & _
    '"    COALESCE(t4e.DEPOSIT_AMOUNT,'0') AS AKI1," & vbCrLf & _
    '"    COALESCE(t4f.DEPOSIT_AMOUNT,'0') AS AKI2," & vbCrLf & _
    '"    COALESCE(t4g.DEPOSIT_AMOUNT,'0') AS AKI3," & vbCrLf & _
    '"    COALESCE(t4h.DEPOSIT_AMOUNT,'0') AS AKI4," & vbCrLf & _
    '"    COALESCE(t5a.DEPOSIT_AMOUNT,'0') AS SUICA_GENKIN," & vbCrLf & _
    '"    COALESCE(t5b.DEPOSIT_AMOUNT,'0') AS SUICA_COIN_LOCKER," & vbCrLf & _
    '"    COALESCE(t5c.DEPOSIT_AMOUNT,'0') AS SONOTA" & vbCrLf & _
    '"FROM" & vbCrLf & _
    '"    YOYAKU_TBL t1" & vbCrLf & _
    '"    LEFT JOIN" & vbCrLf & _
    '"        YDT_TBL t2" & vbCrLf & _
    '"    ON  t1.YOYAKU_NO = t2.YOYAKU_NO" & vbCrLf & _
    '"    LEFT JOIN" & vbCrLf & _
    '"       (" & vbCrLf & _
    '"            SELECT" & vbCrLf & _
    '"                YOYAKU_NO," & vbCrLf & _
    '"                YOYAKU_DT," & vbCrLf & _
    '"                SUM(FUTAI_KIN) AS FUTAI_KIN," & vbCrLf & _
    '"                SUM(FUTAI_CHOSEI) AS FUTAI_CHOSEI" & vbCrLf & _
    '"            FROM" & vbCrLf & _
    '"                FRIYO_MEISAI_TBL" & vbCrLf & _
    '"            GROUP BY" & vbCrLf & _
    '"                YOYAKU_NO," & vbCrLf & _
    '"                YOYAKU_DT" & vbCrLf & _
    '"        ) t3" & vbCrLf & _
    '"    ON  t1.YOYAKU_NO = t3.YOYAKU_NO" & vbCrLf & _
    '"        AND t2.YOYAKU_DT = t3.YOYAKU_DT" & vbCrLf & _
    '"    LEFT JOIN" & vbCrLf & _
    '"        ALSOK_DEPOSIT_TBL t4a" & vbCrLf & _
    '"    ON  t1.YOYAKU_NO = t4a.YOYAKU_NO" & vbCrLf & _
    '"        AND t2.YOYAKU_DT = t4a.DEPOSIT_DT" & vbCrLf & _
    '"        AND t4a.DEPOSIT_KBN = '1'" & vbCrLf & _
    '"        AND t4a.REGISTER_CD= '001'" & vbCrLf & _
    '"    LEFT JOIN" & vbCrLf & _
    '"        ALSOK_DEPOSIT_TBL t4b" & vbCrLf & _
    '"    ON  t1.YOYAKU_NO = t4b.YOYAKU_NO" & vbCrLf & _
    '"        AND t2.YOYAKU_DT = t4b.DEPOSIT_DT" & vbCrLf & _
    '"        AND t4b.DEPOSIT_KBN = '1'" & vbCrLf & _
    '"        AND t4b.REGISTER_CD= '002'" & vbCrLf & _
    '"    LEFT JOIN" & vbCrLf & _
    '"        ALSOK_DEPOSIT_TBL t4c" & vbCrLf & _
    '"    ON  t1.YOYAKU_NO = t4c.YOYAKU_NO" & vbCrLf & _
    '"        AND t2.YOYAKU_DT = t4c.DEPOSIT_DT" & vbCrLf & _
    '"        AND t4c.DEPOSIT_KBN = '1'" & vbCrLf & _
    '"        AND t4c.REGISTER_CD= '003'" & vbCrLf & _
    '"    LEFT JOIN" & vbCrLf & _
    '"        ALSOK_DEPOSIT_TBL t4d" & vbCrLf & _
    '"    ON  t1.YOYAKU_NO = t4d.YOYAKU_NO" & vbCrLf & _
    '"        AND t2.YOYAKU_DT = t4d.DEPOSIT_DT" & vbCrLf & _
    '"        AND t4d.DEPOSIT_KBN = '1'" & vbCrLf & _
    '"        AND t4d.REGISTER_CD= '004'" & vbCrLf & _
    '"    LEFT JOIN" & vbCrLf & _
    '"        ALSOK_DEPOSIT_TBL t4e" & vbCrLf & _
    '"    ON  t1.YOYAKU_NO = t4e.YOYAKU_NO" & vbCrLf & _
    '"        AND t2.YOYAKU_DT = t4e.DEPOSIT_DT" & vbCrLf & _
    '"        AND t4e.DEPOSIT_KBN = '1'" & vbCrLf & _
    '"        AND t4e.REGISTER_CD= '006'" & vbCrLf & _
    '"    LEFT JOIN" & vbCrLf & _
    '"        ALSOK_DEPOSIT_TBL t4f" & vbCrLf & _
    '"    ON  t1.YOYAKU_NO = t4f.YOYAKU_NO" & vbCrLf & _
    '"        AND t2.YOYAKU_DT = t4f.DEPOSIT_DT" & vbCrLf & _
    '"        AND t4f.DEPOSIT_KBN = '1'" & vbCrLf & _
    '"        AND t4f.REGISTER_CD= '007'" & vbCrLf & _
    '"    LEFT JOIN" & vbCrLf & _
    '"        ALSOK_DEPOSIT_TBL t4g" & vbCrLf & _
    '"    ON  t1.YOYAKU_NO = t4g.YOYAKU_NO" & vbCrLf & _
    '"        AND t2.YOYAKU_DT = t4g.DEPOSIT_DT" & vbCrLf & _
    '"        AND t4g.DEPOSIT_KBN = '1'" & vbCrLf & _
    '"        AND t4g.REGISTER_CD= '008'" & vbCrLf & _
    '"    LEFT JOIN" & vbCrLf & _
    '"        ALSOK_DEPOSIT_TBL t4h" & vbCrLf & _
    '"    ON  t1.YOYAKU_NO = t4h.YOYAKU_NO" & vbCrLf & _
    '"        AND t2.YOYAKU_DT = t4h.DEPOSIT_DT" & vbCrLf & _
    '"        AND t4h.DEPOSIT_KBN = '1'" & vbCrLf & _
    '"        AND t4h.REGISTER_CD= '009'" & vbCrLf & _
    '"    LEFT JOIN" & vbCrLf & _
    '"        ALSOK_DEPOSIT_TBL t5a" & vbCrLf & _
    '"    ON  t1.YOYAKU_NO = t5a.YOYAKU_NO" & vbCrLf & _
    '"        AND t2.YOYAKU_DT = t5a.DEPOSIT_DT" & vbCrLf & _
    '"        AND t5a.DEPOSIT_KBN = '2'" & vbCrLf & _
    '"        AND t5a.REGISTER_CD= '900'" & vbCrLf & _
    '"    LEFT JOIN" & vbCrLf & _
    '"        ALSOK_DEPOSIT_TBL t5b" & vbCrLf & _
    '"    ON  t1.YOYAKU_NO = t5b.YOYAKU_NO" & vbCrLf & _
    '"        AND t2.YOYAKU_DT = t5b.DEPOSIT_DT" & vbCrLf & _
    '"        AND t5b.DEPOSIT_KBN = '2'" & vbCrLf & _
    '"        AND t5b.REGISTER_CD= '901'" & vbCrLf & _
    '"    LEFT JOIN" & vbCrLf & _
    '"        (" & vbCrLf & _
    '"            SELECT" & vbCrLf & _
    '"                YOYAKU_NO," & vbCrLf & _
    '"                DEPOSIT_DT," & vbCrLf & _
    '"                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf & _
    '"            FROM" & vbCrLf & _
    '"                ALSOK_DEPOSIT_TBL" & vbCrLf & _
    '"            WHERE" & vbCrLf & _
    '"                DEPOSIT_KBN = '2'" & vbCrLf & _
    '"            AND REGISTER_CD BETWEEN '902' AND '950'" & vbCrLf & _
    '"            GROUP BY" & vbCrLf & _
    '"                YOYAKU_NO," & vbCrLf & _
    '"                DEPOSIT_DT" & vbCrLf & _
    '"        ) t5c" & vbCrLf & _
    '"    ON  t1.YOYAKU_NO = t5c.YOYAKU_NO" & vbCrLf & _
    '"        AND t2.YOYAKU_DT = t5c.DEPOSIT_DT" & vbCrLf & _
    '"    WHERE " & vbCrLf & _
    '"        t2.SHISETU_KBN = '1'" & vbCrLf

    ' 2017.04.17 e.watanabe UPD START↓ シアター利用形状の修正（1:着席、2:スタンディング）
    '  Private strSelectDayUriageTheatreSQL As String = _
    '            "SELECT" & vbCrLf & _
    '            "    t1.YOYAKU_NO," & vbCrLf & _
    '            "    t2.YOYAKU_DT," & vbCrLf & _
    '            "    t1.SAIJI_NM," & vbCrLf & _
    '            "    t1.RIYO_NM," & vbCrLf & _
    '            "    CASE t1.KASHI_KIND" & vbCrLf & _
    '            "        WHEN '0' THEN '未定'" & vbCrLf & _
    '            "        WHEN '1' THEN '一般貸出'" & vbCrLf & _
    '                            "        WHEN '2' THEN '社内利用'" & vbCrLf & _
    '            "        WHEN '3' THEN '特例'" & vbCrLf & _
    '            "    END KASHI_KIND," & vbCrLf & _
    '            "    CASE t1.RIYO_TYPE" & vbCrLf & _
    '            "        WHEN '0' THEN '未定'" & vbCrLf & _
    '            "        WHEN '1' THEN 'スタンディング'" & vbCrLf & _
    '            "        WHEN '2' THEN '着席'" & vbCrLf & _
    '            "        WHEN '3' THEN '変則'" & vbCrLf & _
    '            "        WHEN '4' THEN '催事'" & vbCrLf & _
    '            "    END RIYO_TYPE," & vbCrLf & _
    '            "    CASE t1.SAIJI_BUNRUI" & vbCrLf & _
    '                            "        WHEN '0' THEN '未定'" & vbCrLf & _
    '            "        WHEN '1' THEN '音楽'" & vbCrLf & _
    '            "        WHEN '2' THEN '演劇'" & vbCrLf & _
    '            "        WHEN '3' THEN '演芸'" & vbCrLf & _
    '            "        WHEN '4' THEN 'ビジネス'" & vbCrLf & _
    '            "        WHEN '5' THEN '試写会・映画'" & vbCrLf & _
    '            "        WHEN '6' THEN 'その他'" & vbCrLf & _
    '            "    END SAIJI_BUNRUI," & vbCrLf & _
    '            "    COALESCE(t2.RIYO_KIN,'0') AS RIYO_KIN," & vbCrLf & _
    '            "    COALESCE(t3.FUTAI_KIN,'0') AS FUTAI_KIN," & vbCrLf & _
    '            "    COALESCE(t3.FUTAI_CHOSEI,'0') AS FUTAI_CHOSEI," & vbCrLf & _
    '            "    COALESCE(t4a.DEPOSIT_AMOUNT,'0') AS DRINK_GENKIN," & vbCrLf & _
    '            "    COALESCE(t4b.DEPOSIT_AMOUNT,'0') AS ONE_DRINK," & vbCrLf & _
    '            "    COALESCE(t4c.DEPOSIT_AMOUNT,'0') AS COIN_LOCKER," & vbCrLf & _
    '            "    COALESCE(t4d.DEPOSIT_AMOUNT,'0') AS ZATSU_SHUNYU," & vbCrLf & _
    '            "    COALESCE(t4e.DEPOSIT_AMOUNT,'0') AS AKI1," & vbCrLf & _
    '            "    COALESCE(t4f.DEPOSIT_AMOUNT,'0') AS AKI2," & vbCrLf & _
    '            "    COALESCE(t4g.DEPOSIT_AMOUNT,'0') AS AKI3," & vbCrLf & _
    '            "    COALESCE(t4h.DEPOSIT_AMOUNT,'0') AS AKI4," & vbCrLf & _
    '            "    COALESCE(t5a.DEPOSIT_AMOUNT,'0') AS SUICA_GENKIN," & vbCrLf & _
    '            "    COALESCE(t5b.DEPOSIT_AMOUNT,'0') AS SUICA_COIN_LOCKER," & vbCrLf & _
    '            "    COALESCE(t5c.DEPOSIT_AMOUNT,'0') AS SONOTA " & vbCrLf & _
    '            "FROM" & vbCrLf & _
    '            "    YOYAKU_TBL t1" & vbCrLf & _
    '            "    LEFT JOIN" & vbCrLf & _
    '                            "        YDT_TBL t2" & vbCrLf & _
    '            "    ON  t1.YOYAKU_NO = t2.YOYAKU_NO" & vbCrLf & _
    '            "    LEFT JOIN" & vbCrLf & _
    '            "       (" & vbCrLf & _
    '            "            SELECT" & vbCrLf & _
    '            "                YOYAKU_NO," & vbCrLf & _
    '            "                YOYAKU_DT," & vbCrLf & _
    '            "                SUM(FUTAI_KIN) AS FUTAI_KIN," & vbCrLf & _
    '            "                SUM(FUTAI_CHOSEI) AS FUTAI_CHOSEI" & vbCrLf & _
    '            "            FROM" & vbCrLf & _
    '            "                FRIYO_MEISAI_TBL" & vbCrLf & _
    '                            "            GROUP BY" & vbCrLf & _
    '            "                YOYAKU_NO," & vbCrLf & _
    '            "                YOYAKU_DT" & vbCrLf & _
    '            "        ) t3" & vbCrLf & _
    '            "    ON  t1.YOYAKU_NO = t3.YOYAKU_NO" & vbCrLf & _
    '            "        AND t2.YOYAKU_DT = t3.YOYAKU_DT" & vbCrLf & _
    '            "    LEFT JOIN" & vbCrLf & _
    '            "        (" & vbCrLf & _
    '            "            SELECT" & vbCrLf & _
    '            "                YOYAKU_NO," & vbCrLf & _
    '            "                DEPOSIT_DT," & vbCrLf & _
    '            "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf & _
    '            "            FROM" & vbCrLf & _
    '            "                ALSOK_DEPOSIT_TBL" & vbCrLf & _
    '            "            WHERE" & vbCrLf & _
    '            "                DEPOSIT_KBN = '1'" & vbCrLf & _
    '            "            AND REGISTER_CD= '001'" & vbCrLf & _
    '            "            GROUP BY" & vbCrLf & _
    '            "                YOYAKU_NO," & vbCrLf & _
    '            "                DEPOSIT_DT" & vbCrLf & _
    '            "        ) t4a" & vbCrLf & _
    '            "    ON  t1.YOYAKU_NO = t4a.YOYAKU_NO" & vbCrLf & _
    '            "        AND t2.YOYAKU_DT = t4a.DEPOSIT_DT" & vbCrLf & _
    '            "    LEFT JOIN" & vbCrLf & _
    '            "        (" & vbCrLf & _
    '            "            SELECT" & vbCrLf & _
    '            "                YOYAKU_NO," & vbCrLf & _
    '            "                DEPOSIT_DT," & vbCrLf & _
    '            "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf & _
    '            "            FROM" & vbCrLf & _
    '            "                ALSOK_DEPOSIT_TBL" & vbCrLf & _
    '            "            WHERE" & vbCrLf & _
    '            "                DEPOSIT_KBN = '1'" & vbCrLf & _
    '            "            AND REGISTER_CD= '002'" & vbCrLf & _
    '            "            GROUP BY" & vbCrLf & _
    '            "                YOYAKU_NO," & vbCrLf & _
    '            "                DEPOSIT_DT" & vbCrLf & _
    '            "        ) t4b" & vbCrLf & _
    '            "    ON  t1.YOYAKU_NO = t4b.YOYAKU_NO" & vbCrLf & _
    '            "        AND t2.YOYAKU_DT = t4b.DEPOSIT_DT" & vbCrLf & _
    '            "    LEFT JOIN" & vbCrLf & _
    '            "        (" & vbCrLf & _
    '            "            SELECT" & vbCrLf & _
    '            "                YOYAKU_NO," & vbCrLf & _
    '            "                DEPOSIT_DT," & vbCrLf & _
    '            "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf & _
    '            "            FROM" & vbCrLf & _
    '            "                ALSOK_DEPOSIT_TBL" & vbCrLf & _
    '            "            WHERE" & vbCrLf & _
    '            "                DEPOSIT_KBN = '1'" & vbCrLf & _
    '            "            AND REGISTER_CD= '003'" & vbCrLf & _
    '            "            GROUP BY" & vbCrLf & _
    '            "                YOYAKU_NO," & vbCrLf & _
    '            "                DEPOSIT_DT" & vbCrLf & _
    '            "        ) t4c" & vbCrLf & _
    '            "    ON  t1.YOYAKU_NO = t4c.YOYAKU_NO" & vbCrLf & _
    '            "        AND t2.YOYAKU_DT = t4c.DEPOSIT_DT" & vbCrLf & _
    '            "    LEFT JOIN" & vbCrLf & _
    '            "        (" & vbCrLf & _
    '            "            SELECT" & vbCrLf & _
    '            "                YOYAKU_NO," & vbCrLf & _
    '            "                DEPOSIT_DT," & vbCrLf & _
    '            "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf & _
    '            "            FROM" & vbCrLf & _
    '            "                ALSOK_DEPOSIT_TBL" & vbCrLf & _
    '            "            WHERE" & vbCrLf & _
    '            "                DEPOSIT_KBN = '1'" & vbCrLf & _
    '            "            AND REGISTER_CD= '004'" & vbCrLf & _
    '            "            GROUP BY" & vbCrLf & _
    '            "                YOYAKU_NO," & vbCrLf & _
    '            "                DEPOSIT_DT" & vbCrLf & _
    '            "        ) t4d" & vbCrLf & _
    '            "    ON  t1.YOYAKU_NO = t4d.YOYAKU_NO" & vbCrLf & _
    '            "        AND t2.YOYAKU_DT = t4d.DEPOSIT_DT" & vbCrLf & _
    '            "    LEFT JOIN" & vbCrLf & _
    '            "        (" & vbCrLf & _
    '            "            SELECT" & vbCrLf & _
    '            "                YOYAKU_NO," & vbCrLf & _
    '            "                DEPOSIT_DT," & vbCrLf & _
    '            "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf & _
    '            "            FROM" & vbCrLf & _
    '            "                ALSOK_DEPOSIT_TBL" & vbCrLf & _
    '            "            WHERE" & vbCrLf & _
    '            "                DEPOSIT_KBN = '1'" & vbCrLf & _
    '            "            AND REGISTER_CD= '006'" & vbCrLf & _
    '            "            GROUP BY" & vbCrLf & _
    '            "                YOYAKU_NO," & vbCrLf & _
    '            "                DEPOSIT_DT" & vbCrLf & _
    '            "        ) t4e" & vbCrLf & _
    '            "    ON  t1.YOYAKU_NO = t4e.YOYAKU_NO" & vbCrLf & _
    '            "        AND t2.YOYAKU_DT = t4e.DEPOSIT_DT" & vbCrLf & _
    '            "    LEFT JOIN" & vbCrLf & _
    '            "        (" & vbCrLf & _
    '            "            SELECT" & vbCrLf & _
    '            "                YOYAKU_NO," & vbCrLf & _
    '            "                DEPOSIT_DT," & vbCrLf & _
    '            "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf & _
    '            "            FROM" & vbCrLf & _
    '            "                ALSOK_DEPOSIT_TBL" & vbCrLf & _
    '            "            WHERE" & vbCrLf & _
    '            "                DEPOSIT_KBN = '1'" & vbCrLf & _
    '            "            AND REGISTER_CD= '007'" & vbCrLf & _
    '            "            GROUP BY" & vbCrLf & _
    '            "                YOYAKU_NO," & vbCrLf & _
    '            "                DEPOSIT_DT" & vbCrLf & _
    '            "        ) t4f" & vbCrLf & _
    '            "    ON  t1.YOYAKU_NO = t4f.YOYAKU_NO" & vbCrLf & _
    '            "        AND t2.YOYAKU_DT = t4f.DEPOSIT_DT" & vbCrLf & _
    '            "    LEFT JOIN" & vbCrLf & _
    '            "        (" & vbCrLf & _
    '            "            SELECT" & vbCrLf & _
    '            "                YOYAKU_NO," & vbCrLf & _
    '            "                DEPOSIT_DT," & vbCrLf & _
    '            "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf & _
    '            "            FROM" & vbCrLf & _
    '            "                ALSOK_DEPOSIT_TBL" & vbCrLf & _
    '            "            WHERE" & vbCrLf & _
    '            "                DEPOSIT_KBN = '1'" & vbCrLf & _
    '            "            AND REGISTER_CD= '008'" & vbCrLf & _
    '            "            GROUP BY" & vbCrLf & _
    '            "                YOYAKU_NO," & vbCrLf & _
    '            "                DEPOSIT_DT" & vbCrLf & _
    '            "       ) t4g" & vbCrLf & _
    '            "    ON  t1.YOYAKU_NO = t4g.YOYAKU_NO" & vbCrLf & _
    '            "        AND t2.YOYAKU_DT = t4g.DEPOSIT_DT" & vbCrLf & _
    '            "    LEFT JOIN" & vbCrLf & _
    '            "        (" & vbCrLf & _
    '            "            SELECT" & vbCrLf & _
    '            "                YOYAKU_NO," & vbCrLf & _
    '            "                DEPOSIT_DT," & vbCrLf & _
    '            "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf & _
    '            "            FROM" & vbCrLf & _
    '            "                ALSOK_DEPOSIT_TBL" & vbCrLf & _
    '            "            WHERE" & vbCrLf & _
    '            "                DEPOSIT_KBN = '1'" & vbCrLf & _
    '            "            AND REGISTER_CD= '009'" & vbCrLf & _
    '            "            GROUP BY" & vbCrLf & _
    '            "                YOYAKU_NO," & vbCrLf & _
    '            "                DEPOSIT_DT" & vbCrLf & _
    '            "        ) t4h" & vbCrLf & _
    '            "    ON  t1.YOYAKU_NO = t4h.YOYAKU_NO" & vbCrLf & _
    '            "        AND t2.YOYAKU_DT = t4h.DEPOSIT_DT" & vbCrLf & _
    '            "    LEFT JOIN" & vbCrLf & _
    '            "        (" & vbCrLf & _
    '            "            SELECT" & vbCrLf & _
    '            "                YOYAKU_NO," & vbCrLf & _
    '            "                DEPOSIT_DT," & vbCrLf & _
    '            "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf & _
    '            "            FROM" & vbCrLf & _
    '            "                ALSOK_DEPOSIT_TBL" & vbCrLf & _
    '            "            WHERE" & vbCrLf & _
    '            "                DEPOSIT_KBN = '2'" & vbCrLf & _
    '            "            AND REGISTER_CD= '900'" & vbCrLf & _
    '            "            GROUP BY" & vbCrLf & _
    '            "                YOYAKU_NO," & vbCrLf & _
    '            "                DEPOSIT_DT" & vbCrLf & _
    '            "        ) t5a" & vbCrLf & _
    '            "    ON  t1.YOYAKU_NO = t5a.YOYAKU_NO" & vbCrLf & _
    '            "        AND t2.YOYAKU_DT = t5a.DEPOSIT_DT" & vbCrLf & _
    '            "    LEFT JOIN" & vbCrLf & _
    '            "        (" & vbCrLf & _
    '            "            SELECT" & vbCrLf & _
    '            "                YOYAKU_NO," & vbCrLf & _
    '            "                DEPOSIT_DT," & vbCrLf & _
    '            "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf & _
    '            "            FROM" & vbCrLf & _
    '            "                ALSOK_DEPOSIT_TBL" & vbCrLf & _
    '            "            WHERE" & vbCrLf & _
    '            "                DEPOSIT_KBN = '2'" & vbCrLf & _
    '            "            AND REGISTER_CD= '901'" & vbCrLf & _
    '            "            GROUP BY" & vbCrLf & _
    '            "                YOYAKU_NO," & vbCrLf & _
    '            "                DEPOSIT_DT" & vbCrLf & _
    '            "        ) t5b" & vbCrLf & _
    '            "    ON  t1.YOYAKU_NO = t5b.YOYAKU_NO" & vbCrLf & _
    '            "        AND t2.YOYAKU_DT = t5b.DEPOSIT_DT" & vbCrLf & _
    '            "    LEFT JOIN" & vbCrLf & _
    '            "        (" & vbCrLf & _
    '            "            SELECT" & vbCrLf & _
    '            "                YOYAKU_NO," & vbCrLf & _
    '            "                DEPOSIT_DT," & vbCrLf & _
    '            "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf & _
    '            "            FROM" & vbCrLf & _
    '            "                ALSOK_DEPOSIT_TBL" & vbCrLf & _
    '            "            WHERE" & vbCrLf & _
    '            "                DEPOSIT_KBN = '2'" & vbCrLf & _
    '            "            AND REGISTER_CD BETWEEN '902' AND '950'" & vbCrLf & _
    '            "            GROUP BY" & vbCrLf & _
    '            "                YOYAKU_NO," & vbCrLf & _
    '            "                DEPOSIT_DT" & vbCrLf & _
    '            "        ) t5c" & vbCrLf & _
    '            "    ON  t1.YOYAKU_NO = t5c.YOYAKU_NO" & vbCrLf & _
    '            "        AND t2.YOYAKU_DT = t5c.DEPOSIT_DT" & vbCrLf & _
    '            "    WHERE " & vbCrLf & _
    '            "        t2.SHISETU_KBN = '1'" & vbCrLf
    ' 2016.02.15 UPD END↑ h.hagiwara レジごとに金額集計を行い連結するよう対応
    Private strSelectDayUriageTheatreSQL As String = _
                "SELECT" & vbCrLf & _
                "    t1.YOYAKU_NO," & vbCrLf & _
                "    t2.YOYAKU_DT," & vbCrLf & _
                "    t1.SAIJI_NM," & vbCrLf & _
                "    t1.RIYO_NM," & vbCrLf & _
                "    CASE t1.KASHI_KIND" & vbCrLf & _
                "        WHEN '0' THEN '未定'" & vbCrLf & _
                "        WHEN '1' THEN '一般貸出'" & vbCrLf & _
                                "        WHEN '2' THEN '社内利用'" & vbCrLf & _
                "        WHEN '3' THEN '特例'" & vbCrLf & _
                "    END KASHI_KIND," & vbCrLf & _
                "    CASE t1.RIYO_TYPE" & vbCrLf & _
                "        WHEN '0' THEN '未定'" & vbCrLf & _
                "        WHEN '1' THEN '着席'" & vbCrLf & _
                "        WHEN '2' THEN 'スタンディング'" & vbCrLf & _
                "        WHEN '3' THEN '変則'" & vbCrLf & _
                "        WHEN '4' THEN '催事'" & vbCrLf & _
                "    END RIYO_TYPE," & vbCrLf & _
                "    CASE t1.SAIJI_BUNRUI" & vbCrLf & _
                                "        WHEN '0' THEN '未定'" & vbCrLf & _
                "        WHEN '1' THEN '音楽'" & vbCrLf & _
                "        WHEN '2' THEN '演劇'" & vbCrLf & _
                "        WHEN '3' THEN '演芸'" & vbCrLf & _
                "        WHEN '4' THEN 'ビジネス'" & vbCrLf & _
                "        WHEN '5' THEN '試写会・映画'" & vbCrLf & _
                "        WHEN '6' THEN 'その他'" & vbCrLf & _
                "    END SAIJI_BUNRUI," & vbCrLf & _
                "    COALESCE(t2.RIYO_KIN,'0') AS RIYO_KIN," & vbCrLf & _
                "    COALESCE(t3.FUTAI_KIN,'0') AS FUTAI_KIN," & vbCrLf & _
                "    COALESCE(t3.FUTAI_CHOSEI,'0') AS FUTAI_CHOSEI," & vbCrLf & _
                "    COALESCE(t4a.DEPOSIT_AMOUNT,'0') AS DRINK_GENKIN," & vbCrLf & _
                "    COALESCE(t4b.DEPOSIT_AMOUNT,'0') AS ONE_DRINK," & vbCrLf & _
                "    COALESCE(t4c.DEPOSIT_AMOUNT,'0') AS COIN_LOCKER," & vbCrLf & _
                "    COALESCE(t4d.DEPOSIT_AMOUNT,'0') AS ZATSU_SHUNYU," & vbCrLf & _
                "    COALESCE(t4e.DEPOSIT_AMOUNT,'0') AS AKI1," & vbCrLf & _
                "    COALESCE(t4f.DEPOSIT_AMOUNT,'0') AS AKI2," & vbCrLf & _
                "    COALESCE(t4g.DEPOSIT_AMOUNT,'0') AS AKI3," & vbCrLf & _
                "    COALESCE(t4h.DEPOSIT_AMOUNT,'0') AS AKI4," & vbCrLf & _
                "    COALESCE(t5a.DEPOSIT_AMOUNT,'0') AS SUICA_GENKIN," & vbCrLf & _
                "    COALESCE(t5b.DEPOSIT_AMOUNT,'0') AS SUICA_COIN_LOCKER," & vbCrLf & _
                "    COALESCE(t5c.DEPOSIT_AMOUNT,'0') AS SONOTA " & vbCrLf & _
                "FROM" & vbCrLf & _
                "    YOYAKU_TBL t1" & vbCrLf & _
                "    LEFT JOIN" & vbCrLf & _
                                "        YDT_TBL t2" & vbCrLf & _
                "    ON  t1.YOYAKU_NO = t2.YOYAKU_NO" & vbCrLf & _
                "    LEFT JOIN" & vbCrLf & _
                "       (" & vbCrLf & _
                "            SELECT" & vbCrLf & _
                "                YOYAKU_NO," & vbCrLf & _
                "                YOYAKU_DT," & vbCrLf & _
                "                SUM(FUTAI_KIN) AS FUTAI_KIN," & vbCrLf & _
                "                SUM(FUTAI_CHOSEI) AS FUTAI_CHOSEI" & vbCrLf & _
                "            FROM" & vbCrLf & _
                "                FRIYO_MEISAI_TBL" & vbCrLf & _
                                "            GROUP BY" & vbCrLf & _
                "                YOYAKU_NO," & vbCrLf & _
                "                YOYAKU_DT" & vbCrLf & _
                "        ) t3" & vbCrLf & _
                "    ON  t1.YOYAKU_NO = t3.YOYAKU_NO" & vbCrLf & _
                "        AND t2.YOYAKU_DT = t3.YOYAKU_DT" & vbCrLf & _
                "    LEFT JOIN" & vbCrLf & _
                "        (" & vbCrLf & _
                "            SELECT" & vbCrLf & _
                "                YOYAKU_NO," & vbCrLf & _
                "                DEPOSIT_DT," & vbCrLf & _
                "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf & _
                "            FROM" & vbCrLf & _
                "                ALSOK_DEPOSIT_TBL" & vbCrLf & _
                "            WHERE" & vbCrLf & _
                "                DEPOSIT_KBN = '1'" & vbCrLf & _
                "            AND REGISTER_CD= '001'" & vbCrLf & _
                "            GROUP BY" & vbCrLf & _
                "                YOYAKU_NO," & vbCrLf & _
                "                DEPOSIT_DT" & vbCrLf & _
                "        ) t4a" & vbCrLf & _
                "    ON  t1.YOYAKU_NO = t4a.YOYAKU_NO" & vbCrLf & _
                "        AND t2.YOYAKU_DT = t4a.DEPOSIT_DT" & vbCrLf & _
                "    LEFT JOIN" & vbCrLf & _
                "        (" & vbCrLf & _
                "            SELECT" & vbCrLf & _
                "                YOYAKU_NO," & vbCrLf & _
                "                DEPOSIT_DT," & vbCrLf & _
                "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf & _
                "            FROM" & vbCrLf & _
                "                ALSOK_DEPOSIT_TBL" & vbCrLf & _
                "            WHERE" & vbCrLf & _
                "                DEPOSIT_KBN = '1'" & vbCrLf & _
                "            AND REGISTER_CD= '002'" & vbCrLf & _
                "            GROUP BY" & vbCrLf & _
                "                YOYAKU_NO," & vbCrLf & _
                "                DEPOSIT_DT" & vbCrLf & _
                "        ) t4b" & vbCrLf & _
                "    ON  t1.YOYAKU_NO = t4b.YOYAKU_NO" & vbCrLf & _
                "        AND t2.YOYAKU_DT = t4b.DEPOSIT_DT" & vbCrLf & _
                "    LEFT JOIN" & vbCrLf & _
                "        (" & vbCrLf & _
                "            SELECT" & vbCrLf & _
                "                YOYAKU_NO," & vbCrLf & _
                "                DEPOSIT_DT," & vbCrLf & _
                "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf & _
                "            FROM" & vbCrLf & _
                "                ALSOK_DEPOSIT_TBL" & vbCrLf & _
                "            WHERE" & vbCrLf & _
                "                DEPOSIT_KBN = '1'" & vbCrLf & _
                "            AND REGISTER_CD= '003'" & vbCrLf & _
                "            GROUP BY" & vbCrLf & _
                "                YOYAKU_NO," & vbCrLf & _
                "                DEPOSIT_DT" & vbCrLf & _
                "        ) t4c" & vbCrLf & _
                "    ON  t1.YOYAKU_NO = t4c.YOYAKU_NO" & vbCrLf & _
                "        AND t2.YOYAKU_DT = t4c.DEPOSIT_DT" & vbCrLf & _
                "    LEFT JOIN" & vbCrLf & _
                "        (" & vbCrLf & _
                "            SELECT" & vbCrLf & _
                "                YOYAKU_NO," & vbCrLf & _
                "                DEPOSIT_DT," & vbCrLf & _
                "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf & _
                "            FROM" & vbCrLf & _
                "                ALSOK_DEPOSIT_TBL" & vbCrLf & _
                "            WHERE" & vbCrLf & _
                "                DEPOSIT_KBN = '1'" & vbCrLf & _
                "            AND REGISTER_CD= '004'" & vbCrLf & _
                "            GROUP BY" & vbCrLf & _
                "                YOYAKU_NO," & vbCrLf & _
                "                DEPOSIT_DT" & vbCrLf & _
                "        ) t4d" & vbCrLf & _
                "    ON  t1.YOYAKU_NO = t4d.YOYAKU_NO" & vbCrLf & _
                "        AND t2.YOYAKU_DT = t4d.DEPOSIT_DT" & vbCrLf & _
                "    LEFT JOIN" & vbCrLf & _
                "        (" & vbCrLf & _
                "            SELECT" & vbCrLf & _
                "                YOYAKU_NO," & vbCrLf & _
                "                DEPOSIT_DT," & vbCrLf & _
                "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf & _
                "            FROM" & vbCrLf & _
                "                ALSOK_DEPOSIT_TBL" & vbCrLf & _
                "            WHERE" & vbCrLf & _
                "                DEPOSIT_KBN = '1'" & vbCrLf & _
                "            AND REGISTER_CD= '006'" & vbCrLf & _
                "            GROUP BY" & vbCrLf & _
                "                YOYAKU_NO," & vbCrLf & _
                "                DEPOSIT_DT" & vbCrLf & _
                "        ) t4e" & vbCrLf & _
                "    ON  t1.YOYAKU_NO = t4e.YOYAKU_NO" & vbCrLf & _
                "        AND t2.YOYAKU_DT = t4e.DEPOSIT_DT" & vbCrLf & _
                "    LEFT JOIN" & vbCrLf & _
                "        (" & vbCrLf & _
                "            SELECT" & vbCrLf & _
                "                YOYAKU_NO," & vbCrLf & _
                "                DEPOSIT_DT," & vbCrLf & _
                "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf & _
                "            FROM" & vbCrLf & _
                "                ALSOK_DEPOSIT_TBL" & vbCrLf & _
                "            WHERE" & vbCrLf & _
                "                DEPOSIT_KBN = '1'" & vbCrLf & _
                "            AND REGISTER_CD= '007'" & vbCrLf & _
                "            GROUP BY" & vbCrLf & _
                "                YOYAKU_NO," & vbCrLf & _
                "                DEPOSIT_DT" & vbCrLf & _
                "        ) t4f" & vbCrLf & _
                "    ON  t1.YOYAKU_NO = t4f.YOYAKU_NO" & vbCrLf & _
                "        AND t2.YOYAKU_DT = t4f.DEPOSIT_DT" & vbCrLf & _
                "    LEFT JOIN" & vbCrLf & _
                "        (" & vbCrLf & _
                "            SELECT" & vbCrLf & _
                "                YOYAKU_NO," & vbCrLf & _
                "                DEPOSIT_DT," & vbCrLf & _
                "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf & _
                "            FROM" & vbCrLf & _
                "                ALSOK_DEPOSIT_TBL" & vbCrLf & _
                "            WHERE" & vbCrLf & _
                "                DEPOSIT_KBN = '1'" & vbCrLf & _
                "            AND REGISTER_CD= '008'" & vbCrLf & _
                "            GROUP BY" & vbCrLf & _
                "                YOYAKU_NO," & vbCrLf & _
                "                DEPOSIT_DT" & vbCrLf & _
                "       ) t4g" & vbCrLf & _
                "    ON  t1.YOYAKU_NO = t4g.YOYAKU_NO" & vbCrLf & _
                "        AND t2.YOYAKU_DT = t4g.DEPOSIT_DT" & vbCrLf & _
                "    LEFT JOIN" & vbCrLf & _
                "        (" & vbCrLf & _
                "            SELECT" & vbCrLf & _
                "                YOYAKU_NO," & vbCrLf & _
                "                DEPOSIT_DT," & vbCrLf & _
                "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf & _
                "            FROM" & vbCrLf & _
                "                ALSOK_DEPOSIT_TBL" & vbCrLf & _
                "            WHERE" & vbCrLf & _
                "                DEPOSIT_KBN = '1'" & vbCrLf & _
                "            AND REGISTER_CD= '009'" & vbCrLf & _
                "            GROUP BY" & vbCrLf & _
                "                YOYAKU_NO," & vbCrLf & _
                "                DEPOSIT_DT" & vbCrLf & _
                "        ) t4h" & vbCrLf & _
                "    ON  t1.YOYAKU_NO = t4h.YOYAKU_NO" & vbCrLf & _
                "        AND t2.YOYAKU_DT = t4h.DEPOSIT_DT" & vbCrLf & _
                "    LEFT JOIN" & vbCrLf & _
                "        (" & vbCrLf & _
                "            SELECT" & vbCrLf & _
                "                YOYAKU_NO," & vbCrLf & _
                "                DEPOSIT_DT," & vbCrLf & _
                "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf & _
                "            FROM" & vbCrLf & _
                "                ALSOK_DEPOSIT_TBL" & vbCrLf & _
                "            WHERE" & vbCrLf & _
                "                DEPOSIT_KBN = '2'" & vbCrLf & _
                "            AND REGISTER_CD= '900'" & vbCrLf & _
                "            GROUP BY" & vbCrLf & _
                "                YOYAKU_NO," & vbCrLf & _
                "                DEPOSIT_DT" & vbCrLf & _
                "        ) t5a" & vbCrLf & _
                "    ON  t1.YOYAKU_NO = t5a.YOYAKU_NO" & vbCrLf & _
                "        AND t2.YOYAKU_DT = t5a.DEPOSIT_DT" & vbCrLf & _
                "    LEFT JOIN" & vbCrLf & _
                "        (" & vbCrLf & _
                "            SELECT" & vbCrLf & _
                "                YOYAKU_NO," & vbCrLf & _
                "                DEPOSIT_DT," & vbCrLf & _
                "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf & _
                "            FROM" & vbCrLf & _
                "                ALSOK_DEPOSIT_TBL" & vbCrLf & _
                "            WHERE" & vbCrLf & _
                "                DEPOSIT_KBN = '2'" & vbCrLf & _
                "            AND REGISTER_CD= '901'" & vbCrLf & _
                "            GROUP BY" & vbCrLf & _
                "                YOYAKU_NO," & vbCrLf & _
                "                DEPOSIT_DT" & vbCrLf & _
                "        ) t5b" & vbCrLf & _
                "    ON  t1.YOYAKU_NO = t5b.YOYAKU_NO" & vbCrLf & _
                "        AND t2.YOYAKU_DT = t5b.DEPOSIT_DT" & vbCrLf & _
                "    LEFT JOIN" & vbCrLf & _
                "        (" & vbCrLf & _
                "            SELECT" & vbCrLf & _
                "                YOYAKU_NO," & vbCrLf & _
                "                DEPOSIT_DT," & vbCrLf & _
                "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf & _
                "            FROM" & vbCrLf & _
                "                ALSOK_DEPOSIT_TBL" & vbCrLf & _
                "            WHERE" & vbCrLf & _
                "                DEPOSIT_KBN = '2'" & vbCrLf & _
                "            AND REGISTER_CD BETWEEN '902' AND '950'" & vbCrLf & _
                "            GROUP BY" & vbCrLf & _
                "                YOYAKU_NO," & vbCrLf & _
                "                DEPOSIT_DT" & vbCrLf & _
                "        ) t5c" & vbCrLf & _
                "    ON  t1.YOYAKU_NO = t5c.YOYAKU_NO" & vbCrLf & _
                "        AND t2.YOYAKU_DT = t5c.DEPOSIT_DT" & vbCrLf & _
                "    WHERE " & vbCrLf & _
                "        t2.SHISETU_KBN = '1'" & vbCrLf
    ' 2017.04.17 e.watanabe UPD END↑ シアター利用形状の修正（1:着席、2:スタンディング）

#End Region

#Region "SQL文(請求時の調整額(シアター)取得)"

    'SQL文(請求時の調整額(シアター)取得)
    ' 2017.04.17 e.watanabe UPD START↓ シアター利用形状の修正（1:着席、2:スタンディング）
    '    Private strSelectSeikyuChoseiTheatreSQL As String = _
    '"SELECT" & vbCrLf & _
    '"    t1.YOYAKU_NO," & vbCrLf & _
    '"    t2.YOYAKU_DT," & vbCrLf & _
    '"    t1.SAIJI_NM," & vbCrLf & _
    '"    t1.RIYO_NM," & vbCrLf & _
    '"    CASE t1.KASHI_KIND" & vbCrLf & _
    '"        WHEN '0' THEN '未定'" & vbCrLf & _
    '"        WHEN '1' THEN '一般貸出'" & vbCrLf & _
    '"        WHEN '2' THEN '社内利用'" & vbCrLf & _
    '"        WHEN '3' THEN '特例'" & vbCrLf & _
    '"    END KASHI_KIND," & vbCrLf & _
    '"    CASE t1.RIYO_TYPE" & vbCrLf & _
    '"        WHEN '0' THEN '未定'" & vbCrLf & _
    '"        WHEN '1' THEN 'スタンディング'" & vbCrLf & _
    '"        WHEN '2' THEN '着席'" & vbCrLf & _
    '"        WHEN '3' THEN '変則'" & vbCrLf & _
    '"        WHEN '4' THEN '催事'" & vbCrLf & _
    '"    END RIYO_TYPE," & vbCrLf & _
    '"    CASE t1.SAIJI_BUNRUI" & vbCrLf & _
    '"        WHEN '0' THEN '未定'" & vbCrLf & _
    '"        WHEN '1' THEN '音楽'" & vbCrLf & _
    '"        WHEN '2' THEN '演劇'" & vbCrLf & _
    '"        WHEN '3' THEN '演芸'" & vbCrLf & _
    '"        WHEN '4' THEN 'ビジネス'" & vbCrLf & _
    '"        WHEN '5' THEN '試写会・映画'" & vbCrLf & _
    '"        WHEN '6' THEN 'その他'" & vbCrLf & _
    '"    END SAIJI_BUNRUI," & vbCrLf & _
    '"    t3.SEIKYU_IRAI_NO," & vbCrLf & _
    '"    CASE t3.SEIKYU_NAIYO" & vbCrLf & _
    '"        WHEN '1' THEN '利用料'" & vbCrLf & _
    '"        WHEN '2' THEN '付帯設備'" & vbCrLf & _
    '"        WHEN '3' THEN '利用料+付帯設備'" & vbCrLf & _
    '"        WHEN '4' THEN '還付'" & vbCrLf & _
    '"    END SEIKYU_NAIYO," & vbCrLf & _
    '"    COALESCE(t3.CHOSEI_KIN,'0') AS CHOSEI_KIN" & vbCrLf & _
    '"FROM" & vbCrLf & _
    '"    YOYAKU_TBL t1" & vbCrLf & _
    '"    LEFT JOIN YDT_TBL t2" & vbCrLf & _
    '"    ON  t1.YOYAKU_NO = t2.YOYAKU_NO" & vbCrLf & _
    '"    LEFT JOIN BILLPAY_TBL t3" & vbCrLf & _
    '"    ON  t1.YOYAKU_NO = t3.YOYAKU_NO" & vbCrLf & _
    '"    WHERE " & vbCrLf &
    '"        t2.SHISETU_KBN = '1'" & vbCrLf
    Private strSelectSeikyuChoseiTheatreSQL As String = _
"SELECT" & vbCrLf & _
"    t1.YOYAKU_NO," & vbCrLf & _
"    t2.YOYAKU_DT," & vbCrLf & _
"    t1.SAIJI_NM," & vbCrLf & _
"    t1.RIYO_NM," & vbCrLf & _
"    CASE t1.KASHI_KIND" & vbCrLf & _
"        WHEN '0' THEN '未定'" & vbCrLf & _
"        WHEN '1' THEN '一般貸出'" & vbCrLf & _
"        WHEN '2' THEN '社内利用'" & vbCrLf & _
"        WHEN '3' THEN '特例'" & vbCrLf & _
"    END KASHI_KIND," & vbCrLf & _
"    CASE t1.RIYO_TYPE" & vbCrLf & _
"        WHEN '0' THEN '未定'" & vbCrLf & _
"        WHEN '1' THEN '着席'" & vbCrLf & _
"        WHEN '2' THEN 'スタンディング'" & vbCrLf & _
"        WHEN '3' THEN '変則'" & vbCrLf & _
"        WHEN '4' THEN '催事'" & vbCrLf & _
"    END RIYO_TYPE," & vbCrLf & _
"    CASE t1.SAIJI_BUNRUI" & vbCrLf & _
"        WHEN '0' THEN '未定'" & vbCrLf & _
"        WHEN '1' THEN '音楽'" & vbCrLf & _
"        WHEN '2' THEN '演劇'" & vbCrLf & _
"        WHEN '3' THEN '演芸'" & vbCrLf & _
"        WHEN '4' THEN 'ビジネス'" & vbCrLf & _
"        WHEN '5' THEN '試写会・映画'" & vbCrLf & _
"        WHEN '6' THEN 'その他'" & vbCrLf & _
"    END SAIJI_BUNRUI," & vbCrLf & _
"    t3.SEIKYU_IRAI_NO," & vbCrLf & _
"    CASE t3.SEIKYU_NAIYO" & vbCrLf & _
"        WHEN '1' THEN '利用料'" & vbCrLf & _
"        WHEN '2' THEN '付帯設備'" & vbCrLf & _
"        WHEN '3' THEN '利用料+付帯設備'" & vbCrLf & _
"        WHEN '4' THEN '還付'" & vbCrLf & _
"    END SEIKYU_NAIYO," & vbCrLf & _
"    COALESCE(t3.CHOSEI_KIN,'0') AS CHOSEI_KIN" & vbCrLf & _
"FROM" & vbCrLf & _
"    YOYAKU_TBL t1" & vbCrLf & _
"    LEFT JOIN YDT_TBL t2" & vbCrLf & _
"    ON  t1.YOYAKU_NO = t2.YOYAKU_NO" & vbCrLf & _
"    LEFT JOIN BILLPAY_TBL t3" & vbCrLf & _
"    ON  t1.YOYAKU_NO = t3.YOYAKU_NO" & vbCrLf & _
"    WHERE " & vbCrLf &
"        t2.SHISETU_KBN = '1'" & vbCrLf
    ' 2017.04.17 e.watanabe UPD END↑ シアター利用形状の修正（1:着席、2:スタンディング）

#End Region

#Region "SQL文(日別売上一覧(スタジオ)取得)"

    'SQL文(日別売上一覧(スタジオ)取得)
    ' 2016.02.15 UPD START↓ h.hagiwara レジごとに金額集計を行い連結するよう対応
    'Private strSelectDayUriageStudioSQL As String = _
    '"SELECT" & vbCrLf & _
    '"    t1.YOYAKU_NO," & vbCrLf & _
    '"    t2.YOYAKU_DT," & vbCrLf & _
    '"    t1.SHUTSUEN_NM," & vbCrLf & _
    '"    t1.RIYO_NM," & vbCrLf & _
    '"    CASE t1.KASHI_KIND" & vbCrLf & _
    '"        WHEN '0' THEN '未定'" & vbCrLf & _
    '"        WHEN '1' THEN '一般貸出'" & vbCrLf & _
    '"        WHEN '2' THEN '社内利用'" & vbCrLf & _
    '"    END KASHI_KIND," & vbCrLf & _
    '"    CASE t1.STUDIO_KBN" & vbCrLf & _
    '"        WHEN '1' THEN '201st'" & vbCrLf & _
    '"        WHEN '2' THEN '202st'" & vbCrLf & _
    '"        WHEN '3' THEN 'house lock'" & vbCrLf & _
    '"    END STUDIO_KBN," & vbCrLf & _
    '"    COALESCE(t2.RIYO_KIN,'0') AS RIYO_KIN," & vbCrLf & _
    '"    COALESCE(t3.FUTAI_KIN,'0') AS FUTAI_KIN," & vbCrLf & _
    '"    COALESCE(t3.FUTAI_CHOSEI,'0') AS FUTAI_CHOSEI," & vbCrLf & _
    '"    COALESCE(t4a.DEPOSIT_AMOUNT,'0') AS DRINK_GENKIN," & vbCrLf & _
    '"    COALESCE(t4b.DEPOSIT_AMOUNT,'0') AS ONE_DRINK," & vbCrLf & _
    '"    COALESCE(t4c.DEPOSIT_AMOUNT,'0') AS COIN_LOCKER," & vbCrLf & _
    '"    COALESCE(t4d.DEPOSIT_AMOUNT,'0') AS ZATSU_SHUNYU," & vbCrLf & _
    '"    COALESCE(t4e.DEPOSIT_AMOUNT,'0') AS AKI1," & vbCrLf & _
    '"    COALESCE(t4f.DEPOSIT_AMOUNT,'0') AS AKI2," & vbCrLf & _
    '"    COALESCE(t4g.DEPOSIT_AMOUNT,'0') AS AKI3," & vbCrLf & _
    '"    COALESCE(t4h.DEPOSIT_AMOUNT,'0') AS AKI4," & vbCrLf & _
    '"    COALESCE(t5a.DEPOSIT_AMOUNT,'0') AS SUICA_GENKIN," & vbCrLf & _
    '"    COALESCE(t5b.DEPOSIT_AMOUNT,'0') AS SUICA_COIN_LOCKER," & vbCrLf & _
    '"    COALESCE(t5c.DEPOSIT_AMOUNT,'0') AS SONOTA" & vbCrLf & _
    '"FROM" & vbCrLf & _
    '"    YOYAKU_TBL t1" & vbCrLf & _
    '"    LEFT JOIN" & vbCrLf & _
    '"        YDT_TBL t2" & vbCrLf & _
    '"    ON  t1.YOYAKU_NO = t2.YOYAKU_NO" & vbCrLf & _
    '"    LEFT JOIN" & vbCrLf & _
    '"       (" & vbCrLf & _
    '"            SELECT" & vbCrLf & _
    '"                YOYAKU_NO," & vbCrLf & _
    '"                YOYAKU_DT," & vbCrLf & _
    '"                SUM(FUTAI_KIN) AS FUTAI_KIN," & vbCrLf & _
    '"                SUM(FUTAI_CHOSEI) AS FUTAI_CHOSEI" & vbCrLf & _
    '"            FROM" & vbCrLf & _
    '"                FRIYO_MEISAI_TBL" & vbCrLf & _
    '"            GROUP BY" & vbCrLf & _
    '"                YOYAKU_NO," & vbCrLf & _
    '"                YOYAKU_DT" & vbCrLf & _
    '"        ) t3" & vbCrLf & _
    '"    ON  t1.YOYAKU_NO = t3.YOYAKU_NO" & vbCrLf & _
    '"        AND t2.YOYAKU_DT = t3.YOYAKU_DT" & vbCrLf & _
    '"    LEFT JOIN" & vbCrLf & _
    '"        ALSOK_DEPOSIT_TBL t4a" & vbCrLf & _
    '"    ON  t1.YOYAKU_NO = t4a.YOYAKU_NO" & vbCrLf & _
    '"        AND t2.YOYAKU_DT = t4a.DEPOSIT_DT" & vbCrLf & _
    '"        AND t4a.DEPOSIT_KBN = '1'" & vbCrLf & _
    '"        AND t4a.REGISTER_CD= '001'" & vbCrLf & _
    '"    LEFT JOIN" & vbCrLf & _
    '"        ALSOK_DEPOSIT_TBL t4b" & vbCrLf & _
    '"    ON  t1.YOYAKU_NO = t4b.YOYAKU_NO" & vbCrLf & _
    '"        AND t2.YOYAKU_DT = t4b.DEPOSIT_DT" & vbCrLf & _
    '"        AND t4b.DEPOSIT_KBN = '1'" & vbCrLf & _
    '"        AND t4b.REGISTER_CD= '002'" & vbCrLf & _
    '"    LEFT JOIN" & vbCrLf & _
    '"        ALSOK_DEPOSIT_TBL t4c" & vbCrLf & _
    '"    ON  t1.YOYAKU_NO = t4c.YOYAKU_NO" & vbCrLf & _
    '"        AND t2.YOYAKU_DT = t4c.DEPOSIT_DT" & vbCrLf & _
    '"        AND t4c.DEPOSIT_KBN = '1'" & vbCrLf & _
    '"        AND t4c.REGISTER_CD= '003'" & vbCrLf & _
    '"    LEFT JOIN" & vbCrLf & _
    '"        ALSOK_DEPOSIT_TBL t4d" & vbCrLf & _
    '"    ON  t1.YOYAKU_NO = t4d.YOYAKU_NO" & vbCrLf & _
    '"        AND t2.YOYAKU_DT = t4d.DEPOSIT_DT" & vbCrLf & _
    '"        AND t4d.DEPOSIT_KBN = '1'" & vbCrLf & _
    '"        AND t4d.REGISTER_CD= '004'" & vbCrLf & _
    '"    LEFT JOIN" & vbCrLf & _
    '"        ALSOK_DEPOSIT_TBL t4e" & vbCrLf & _
    '"    ON  t1.YOYAKU_NO = t4e.YOYAKU_NO" & vbCrLf & _
    '"        AND t2.YOYAKU_DT = t4e.DEPOSIT_DT" & vbCrLf & _
    '"        AND t4e.DEPOSIT_KBN = '1'" & vbCrLf & _
    '"        AND t4e.REGISTER_CD= '006'" & vbCrLf & _
    '"    LEFT JOIN" & vbCrLf & _
    '"        ALSOK_DEPOSIT_TBL t4f" & vbCrLf & _
    '"    ON  t1.YOYAKU_NO = t4f.YOYAKU_NO" & vbCrLf & _
    '"        AND t2.YOYAKU_DT = t4f.DEPOSIT_DT" & vbCrLf & _
    '"        AND t4f.DEPOSIT_KBN = '1'" & vbCrLf & _
    '"        AND t4f.REGISTER_CD= '007'" & vbCrLf & _
    '"    LEFT JOIN" & vbCrLf & _
    '"        ALSOK_DEPOSIT_TBL t4g" & vbCrLf & _
    '"    ON  t1.YOYAKU_NO = t4g.YOYAKU_NO" & vbCrLf & _
    '"        AND t2.YOYAKU_DT = t4g.DEPOSIT_DT" & vbCrLf & _
    '"        AND t4g.DEPOSIT_KBN = '1'" & vbCrLf & _
    '"        AND t4g.REGISTER_CD= '008'" & vbCrLf & _
    '"    LEFT JOIN" & vbCrLf & _
    '"        ALSOK_DEPOSIT_TBL t4h" & vbCrLf & _
    '"    ON  t1.YOYAKU_NO = t4h.YOYAKU_NO" & vbCrLf & _
    '"        AND t2.YOYAKU_DT = t4h.DEPOSIT_DT" & vbCrLf & _
    '"        AND t4h.DEPOSIT_KBN = '1'" & vbCrLf & _
    '"        AND t4h.REGISTER_CD= '009'" & vbCrLf & _
    '"    LEFT JOIN" & vbCrLf & _
    '"        ALSOK_DEPOSIT_TBL t5a" & vbCrLf & _
    '"    ON  t1.YOYAKU_NO = t5a.YOYAKU_NO" & vbCrLf & _
    '"        AND t2.YOYAKU_DT = t5a.DEPOSIT_DT" & vbCrLf & _
    '"        AND t5a.DEPOSIT_KBN = '2'" & vbCrLf & _
    '"        AND t5a.REGISTER_CD= '900'" & vbCrLf & _
    '"    LEFT JOIN" & vbCrLf & _
    '"        ALSOK_DEPOSIT_TBL t5b" & vbCrLf & _
    '"    ON  t1.YOYAKU_NO = t5b.YOYAKU_NO" & vbCrLf & _
    '"        AND t2.YOYAKU_DT = t5b.DEPOSIT_DT" & vbCrLf & _
    '"        AND t5b.DEPOSIT_KBN = '2'" & vbCrLf & _
    '"        AND t5b.REGISTER_CD= '901'" & vbCrLf & _
    '"    LEFT JOIN" & vbCrLf & _
    '"        (" & vbCrLf & _
    '"            SELECT" & vbCrLf & _
    '"                YOYAKU_NO," & vbCrLf & _
    '"                DEPOSIT_DT," & vbCrLf & _
    '"                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf & _
    '"            FROM" & vbCrLf & _
    '"                ALSOK_DEPOSIT_TBL" & vbCrLf & _
    '"            WHERE" & vbCrLf & _
    '"                DEPOSIT_KBN = '2'" & vbCrLf & _
    '"            AND REGISTER_CD BETWEEN '902' AND '950'" & vbCrLf & _
    '"            GROUP BY" & vbCrLf & _
    '"                YOYAKU_NO," & vbCrLf & _
    '"                DEPOSIT_DT" & vbCrLf & _
    '"        ) t5c" & vbCrLf & _
    '"    ON  t1.YOYAKU_NO = t5c.YOYAKU_NO" & vbCrLf & _
    '"        AND t2.YOYAKU_DT = t5c.DEPOSIT_DT" & vbCrLf & _
    '"    WHERE " & vbCrLf &
    '"        t2.SHISETU_KBN = '2'" & vbCrLf
    Private strSelectDayUriageStudioSQL As String = _
                "SELECT" & vbCrLf & _
                "    t1.YOYAKU_NO," & vbCrLf & _
                "    t2.YOYAKU_DT," & vbCrLf & _
                "    t1.SHUTSUEN_NM," & vbCrLf & _
                "    t1.RIYO_NM," & vbCrLf & _
                "    CASE t1.KASHI_KIND" & vbCrLf & _
                "        WHEN '0' THEN '未定'" & vbCrLf & _
                "        WHEN '1' THEN '一般貸出'" & vbCrLf & _
                "        WHEN '2' THEN '社内利用'" & vbCrLf & _
                "    END KASHI_KIND," & vbCrLf & _
                "    CASE t1.STUDIO_KBN" & vbCrLf & _
                "        WHEN '1' THEN '201st'" & vbCrLf & _
                "        WHEN '2' THEN '202st'" & vbCrLf & _
                "        WHEN '3' THEN 'house lock'" & vbCrLf & _
                "    END STUDIO_KBN," & vbCrLf & _
                "    COALESCE(t2.RIYO_KIN,'0') AS RIYO_KIN," & vbCrLf & _
                "    COALESCE(t3.FUTAI_KIN,'0') AS FUTAI_KIN," & vbCrLf & _
                "    COALESCE(t3.FUTAI_CHOSEI,'0') AS FUTAI_CHOSEI," & vbCrLf & _
                "    COALESCE(t4a.DEPOSIT_AMOUNT,'0') AS DRINK_GENKIN," & vbCrLf & _
                "    COALESCE(t4b.DEPOSIT_AMOUNT,'0') AS ONE_DRINK," & vbCrLf & _
                "    COALESCE(t4c.DEPOSIT_AMOUNT,'0') AS COIN_LOCKER," & vbCrLf & _
                "    COALESCE(t4d.DEPOSIT_AMOUNT,'0') AS ZATSU_SHUNYU," & vbCrLf & _
                "    COALESCE(t4e.DEPOSIT_AMOUNT,'0') AS AKI1," & vbCrLf & _
                "    COALESCE(t4f.DEPOSIT_AMOUNT,'0') AS AKI2," & vbCrLf & _
                "    COALESCE(t4g.DEPOSIT_AMOUNT,'0') AS AKI3," & vbCrLf & _
                "    COALESCE(t4h.DEPOSIT_AMOUNT,'0') AS AKI4," & vbCrLf & _
                "    COALESCE(t5a.DEPOSIT_AMOUNT,'0') AS SUICA_GENKIN," & vbCrLf & _
                "    COALESCE(t5b.DEPOSIT_AMOUNT,'0') AS SUICA_COIN_LOCKER," & vbCrLf & _
                "    COALESCE(t5c.DEPOSIT_AMOUNT,'0') AS SONOTA" & vbCrLf & _
                "FROM" & vbCrLf & _
                "    YOYAKU_TBL t1" & vbCrLf & _
                "    LEFT JOIN" & vbCrLf & _
                "        YDT_TBL t2" & vbCrLf & _
                "    ON  t1.YOYAKU_NO = t2.YOYAKU_NO" & vbCrLf & _
                "    LEFT JOIN" & vbCrLf & _
                "       (" & vbCrLf & _
                "            SELECT" & vbCrLf & _
                "                YOYAKU_NO," & vbCrLf & _
                "                YOYAKU_DT," & vbCrLf & _
                "                SUM(FUTAI_KIN) AS FUTAI_KIN," & vbCrLf & _
                "                SUM(FUTAI_CHOSEI) AS FUTAI_CHOSEI" & vbCrLf & _
                "            FROM" & vbCrLf & _
                "                FRIYO_MEISAI_TBL" & vbCrLf & _
                "            GROUP BY" & vbCrLf & _
                "                YOYAKU_NO," & vbCrLf & _
                "                YOYAKU_DT" & vbCrLf & _
                "        ) t3" & vbCrLf & _
                "    ON  t1.YOYAKU_NO = t3.YOYAKU_NO" & vbCrLf & _
                "        AND t2.YOYAKU_DT = t3.YOYAKU_DT" & vbCrLf & _
                "    LEFT JOIN" & vbCrLf & _
                "        (" & vbCrLf & _
                "            SELECT" & vbCrLf & _
                "                YOYAKU_NO," & vbCrLf & _
                "                DEPOSIT_DT," & vbCrLf & _
                "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf & _
                "            FROM" & vbCrLf & _
                "                ALSOK_DEPOSIT_TBL" & vbCrLf & _
                "            WHERE" & vbCrLf & _
                "                DEPOSIT_KBN = '1'" & vbCrLf & _
                "            AND REGISTER_CD= '001'" & vbCrLf & _
                "            GROUP BY" & vbCrLf & _
                "                YOYAKU_NO," & vbCrLf & _
                "                DEPOSIT_DT" & vbCrLf & _
                "        ) t4a" & vbCrLf & _
                "    ON  t1.YOYAKU_NO = t4a.YOYAKU_NO" & vbCrLf & _
                "        AND t2.YOYAKU_DT = t4a.DEPOSIT_DT" & vbCrLf & _
                "    LEFT JOIN" & vbCrLf & _
                "        (" & vbCrLf & _
                "            SELECT" & vbCrLf & _
                "                YOYAKU_NO," & vbCrLf & _
                "                DEPOSIT_DT," & vbCrLf & _
                "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf & _
                "            FROM" & vbCrLf & _
                "                ALSOK_DEPOSIT_TBL" & vbCrLf & _
                "            WHERE" & vbCrLf & _
                "                DEPOSIT_KBN = '1'" & vbCrLf & _
                "            AND REGISTER_CD= '002'" & vbCrLf & _
                "            GROUP BY" & vbCrLf & _
                "                YOYAKU_NO," & vbCrLf & _
                "                DEPOSIT_DT" & vbCrLf & _
                "        ) t4b" & vbCrLf & _
                "    ON  t1.YOYAKU_NO = t4b.YOYAKU_NO" & vbCrLf & _
                "        AND t2.YOYAKU_DT = t4b.DEPOSIT_DT" & vbCrLf & _
                "    LEFT JOIN" & vbCrLf & _
                "        (" & vbCrLf & _
                "            SELECT" & vbCrLf & _
                "                YOYAKU_NO," & vbCrLf & _
                "                DEPOSIT_DT," & vbCrLf & _
                "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf & _
                "            FROM" & vbCrLf & _
                "                ALSOK_DEPOSIT_TBL" & vbCrLf & _
                "            WHERE" & vbCrLf & _
                "                DEPOSIT_KBN = '1'" & vbCrLf & _
                "            AND REGISTER_CD= '003'" & vbCrLf & _
                "            GROUP BY" & vbCrLf & _
                "                YOYAKU_NO," & vbCrLf & _
                "                DEPOSIT_DT" & vbCrLf & _
                "        ) t4c" & vbCrLf & _
                "    ON  t1.YOYAKU_NO = t4c.YOYAKU_NO" & vbCrLf & _
                "        AND t2.YOYAKU_DT = t4c.DEPOSIT_DT" & vbCrLf & _
                "    LEFT JOIN" & vbCrLf & _
                "        (" & vbCrLf & _
                "            SELECT" & vbCrLf & _
                "                YOYAKU_NO," & vbCrLf & _
                "                DEPOSIT_DT," & vbCrLf & _
                "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf & _
                "            FROM" & vbCrLf & _
                "                ALSOK_DEPOSIT_TBL" & vbCrLf & _
                "            WHERE" & vbCrLf & _
                "                DEPOSIT_KBN = '1'" & vbCrLf & _
                "            AND REGISTER_CD= '004'" & vbCrLf & _
                "            GROUP BY" & vbCrLf & _
                "                YOYAKU_NO," & vbCrLf & _
                "                DEPOSIT_DT" & vbCrLf & _
                "        ) t4d" & vbCrLf & _
                "    ON  t1.YOYAKU_NO = t4d.YOYAKU_NO" & vbCrLf & _
                "        AND t2.YOYAKU_DT = t4d.DEPOSIT_DT" & vbCrLf & _
                "    LEFT JOIN" & vbCrLf & _
                "        (" & vbCrLf & _
                "            SELECT" & vbCrLf & _
                "                YOYAKU_NO," & vbCrLf & _
                "                DEPOSIT_DT," & vbCrLf & _
                "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf & _
                "            FROM" & vbCrLf & _
                "                ALSOK_DEPOSIT_TBL" & vbCrLf & _
                "            WHERE" & vbCrLf & _
                "                DEPOSIT_KBN = '1'" & vbCrLf & _
                "            AND REGISTER_CD= '006'" & vbCrLf & _
                "            GROUP BY" & vbCrLf & _
                "                YOYAKU_NO," & vbCrLf & _
                "                DEPOSIT_DT" & vbCrLf & _
                "        ) t4e" & vbCrLf & _
                "    ON  t1.YOYAKU_NO = t4e.YOYAKU_NO" & vbCrLf & _
                "        AND t2.YOYAKU_DT = t4e.DEPOSIT_DT" & vbCrLf & _
                "    LEFT JOIN" & vbCrLf & _
                "        (" & vbCrLf & _
                "            SELECT" & vbCrLf & _
                "                YOYAKU_NO," & vbCrLf & _
                "                DEPOSIT_DT," & vbCrLf & _
                "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf & _
                "            FROM" & vbCrLf & _
                "                ALSOK_DEPOSIT_TBL" & vbCrLf & _
                "            WHERE" & vbCrLf & _
                "                DEPOSIT_KBN = '1'" & vbCrLf & _
                "            AND REGISTER_CD= '007'" & vbCrLf & _
                "            GROUP BY" & vbCrLf & _
                "                YOYAKU_NO," & vbCrLf & _
                "                DEPOSIT_DT" & vbCrLf & _
                "        ) t4f" & vbCrLf & _
                "    ON  t1.YOYAKU_NO = t4f.YOYAKU_NO" & vbCrLf & _
                "        AND t2.YOYAKU_DT = t4f.DEPOSIT_DT" & vbCrLf & _
                "    LEFT JOIN" & vbCrLf & _
                "        (" & vbCrLf & _
                "            SELECT" & vbCrLf & _
                "                YOYAKU_NO," & vbCrLf & _
                "                DEPOSIT_DT," & vbCrLf & _
                "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf & _
                "            FROM" & vbCrLf & _
                "                ALSOK_DEPOSIT_TBL" & vbCrLf & _
                "            WHERE" & vbCrLf & _
                "                DEPOSIT_KBN = '1'" & vbCrLf & _
                "            AND REGISTER_CD= '008'" & vbCrLf & _
                "            GROUP BY" & vbCrLf & _
                "                YOYAKU_NO," & vbCrLf & _
                "                DEPOSIT_DT" & vbCrLf & _
                "        ) t4g" & vbCrLf & _
                "    ON  t1.YOYAKU_NO = t4g.YOYAKU_NO" & vbCrLf & _
                "        AND t2.YOYAKU_DT = t4g.DEPOSIT_DT" & vbCrLf & _
                "    LEFT JOIN" & vbCrLf & _
                "        (" & vbCrLf & _
                "            SELECT" & vbCrLf & _
                "                YOYAKU_NO," & vbCrLf & _
                "                DEPOSIT_DT," & vbCrLf & _
                "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf & _
                "            FROM" & vbCrLf & _
                "                ALSOK_DEPOSIT_TBL" & vbCrLf & _
                "            WHERE" & vbCrLf & _
                "                DEPOSIT_KBN = '1'" & vbCrLf & _
                "            AND REGISTER_CD= '009'" & vbCrLf & _
                "            GROUP BY" & vbCrLf & _
                "                YOYAKU_NO," & vbCrLf & _
                "                DEPOSIT_DT" & vbCrLf & _
                "        ) t4h" & vbCrLf & _
                "    ON  t1.YOYAKU_NO = t4h.YOYAKU_NO" & vbCrLf & _
                "        AND t2.YOYAKU_DT = t4h.DEPOSIT_DT" & vbCrLf & _
                "    LEFT JOIN" & vbCrLf & _
                "        (" & vbCrLf & _
                "            SELECT" & vbCrLf & _
                "                YOYAKU_NO," & vbCrLf & _
                "                DEPOSIT_DT," & vbCrLf & _
                "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf & _
                "            FROM" & vbCrLf & _
                "                ALSOK_DEPOSIT_TBL" & vbCrLf & _
                "            WHERE" & vbCrLf & _
                "                DEPOSIT_KBN = '2'" & vbCrLf & _
                "            AND REGISTER_CD= '900'" & vbCrLf & _
                "            GROUP BY" & vbCrLf & _
                "                YOYAKU_NO," & vbCrLf & _
                "                DEPOSIT_DT" & vbCrLf & _
                "        ) t5a" & vbCrLf & _
                "    ON  t1.YOYAKU_NO = t5a.YOYAKU_NO" & vbCrLf & _
                "        AND t2.YOYAKU_DT = t5a.DEPOSIT_DT" & vbCrLf & _
                "    LEFT JOIN" & vbCrLf & _
                "        (" & vbCrLf & _
                "            SELECT" & vbCrLf & _
                "                YOYAKU_NO," & vbCrLf & _
                "                DEPOSIT_DT," & vbCrLf & _
                "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf & _
                "            FROM" & vbCrLf & _
                "                ALSOK_DEPOSIT_TBL" & vbCrLf & _
                "            WHERE" & vbCrLf & _
                "                DEPOSIT_KBN = '2'" & vbCrLf & _
                "            AND REGISTER_CD= '901'" & vbCrLf & _
                "            GROUP BY" & vbCrLf & _
                "                YOYAKU_NO," & vbCrLf & _
                "                DEPOSIT_DT" & vbCrLf & _
                "        ) t5b" & vbCrLf & _
                "    ON  t1.YOYAKU_NO = t5b.YOYAKU_NO" & vbCrLf & _
                "        AND t2.YOYAKU_DT = t5b.DEPOSIT_DT" & vbCrLf & _
                "    LEFT JOIN" & vbCrLf & _
                "        (" & vbCrLf & _
                "            SELECT" & vbCrLf & _
                "                YOYAKU_NO," & vbCrLf & _
                "                DEPOSIT_DT," & vbCrLf & _
                "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf & _
                "            FROM" & vbCrLf & _
                "                ALSOK_DEPOSIT_TBL" & vbCrLf & _
                "            WHERE" & vbCrLf & _
                "                DEPOSIT_KBN = '2'" & vbCrLf & _
                "            AND REGISTER_CD BETWEEN '902' AND '950'" & vbCrLf & _
                "            GROUP BY" & vbCrLf & _
                "                YOYAKU_NO," & vbCrLf & _
                "                DEPOSIT_DT" & vbCrLf & _
                "        ) t5c" & vbCrLf & _
                "    ON  t1.YOYAKU_NO = t5c.YOYAKU_NO" & vbCrLf & _
                "        AND t2.YOYAKU_DT = t5c.DEPOSIT_DT" & vbCrLf & _
                "    WHERE " & vbCrLf &
                "        t2.SHISETU_KBN = '2'" & vbCrLf
    ' 2016.02.15 UPD END↑ h.hagiwara レジごとに金額集計を行い連結するよう対応

#End Region

#Region "SQL文(請求時の調整額(スタジオ)取得)"

    'SQL文(請求時の調整額(スタジオ)取得)
    Private strSelectSeikyuChoseiStudioSQL As String = _
"SELECT" & vbCrLf & _
"    t1.YOYAKU_NO," & vbCrLf & _
"    t2.YOYAKU_DT," & vbCrLf & _
"    t1.SHUTSUEN_NM," & vbCrLf & _
"    t1.RIYO_NM," & vbCrLf & _
"    CASE t1.KASHI_KIND" & vbCrLf & _
"        WHEN '0' THEN '未定'" & vbCrLf & _
"        WHEN '1' THEN '一般貸出'" & vbCrLf & _
"        WHEN '2' THEN '社内利用'" & vbCrLf & _
"    END KASHI_KIND," & vbCrLf & _
"    CASE t1.STUDIO_KBN" & vbCrLf & _
"        WHEN '1' THEN '201st'" & vbCrLf & _
"        WHEN '2' THEN '202st'" & vbCrLf & _
"        WHEN '3' THEN 'house lock'" & vbCrLf & _
"    END STUDIO_KBN," & vbCrLf & _
"    t3.SEIKYU_IRAI_NO," & vbCrLf & _
"    CASE t3.SEIKYU_NAIYO" & vbCrLf & _
"        WHEN '1' THEN '利用料'" & vbCrLf & _
"        WHEN '2' THEN '付帯設備'" & vbCrLf & _
"        WHEN '3' THEN '利用料+付帯設備'" & vbCrLf & _
"        WHEN '4' THEN '還付'" & vbCrLf & _
"    END SEIKYU_NAIYO," & vbCrLf & _
"    COALESCE(t3.CHOSEI_KIN,'0') AS CHOSEI_KIN" & vbCrLf & _
"FROM" & vbCrLf & _
"    YOYAKU_TBL t1" & vbCrLf & _
"    LEFT JOIN YDT_TBL t2" & vbCrLf & _
"    ON  t1.YOYAKU_NO = t2.YOYAKU_NO" & vbCrLf & _
"    LEFT JOIN BILLPAY_TBL t3" & vbCrLf & _
"    ON  t1.YOYAKU_NO = t3.YOYAKU_NO" & vbCrLf & _
"    WHERE " & vbCrLf &
"        t2.SHISETU_KBN = '2'" & vbCrLf

#End Region

    '*****条件追加・変更、アダプタ・パラメータ網羅によるSQL文*****
#Region "日別売上一覧取得"

    ''' <summary>
    ''' 日別売上一覧取得
    ''' </summary>
    ''' <param name="Adapter">アダプタ</param>
    ''' <param name="Cn">コネクション</param>
    ''' <param name="dataEXTZ0103">データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>日別売上一覧を取得するSQLの作成
    ''' <para>作成情報：2015/09/16 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function setSelectDayUriage(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByRef Cn As NpgsqlConnection, _
                                       ByVal dataEXTZ0103 As DataEXTZ0103) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = String.Empty

        Try

            With dataEXTZ0103
                'SQL文(SELECT)
                'データアダプタに、SQLのSELECT文を設定
                If .PropStrShisetsuKbn = SHISETU_KBN_THEATER Then
                    'シアターの場合
                    strSQL = strSelectDayUriageTheatreSQL
                Else
                    'スタジオの場合
                    strSQL = strSelectDayUriageStudioSQL
                End If

                '使用年(From）・使用月(From）が共に入力されている場合
                If .PropStrShiyoNenFrom <> String.Empty AndAlso .PropStrShiyoTsukiFrom <> String.Empty Then
                    strSQL &= "AND TO_DATE(t2.YOYAKU_DT, 'YYYY/MM/DD') >= TO_DATE('" & .PropStrShiyoNenFrom & "/" & .PropStrShiyoTsukiFrom & "/01', 'YYYY/MM/DD') " & vbCrLf
                End If

                '使用年(To）・使用月(To）が共に入力されている場合
                If .PropStrShiyoNenTo <> String.Empty AndAlso .PropStrShiyoTsukiTo <> String.Empty Then
                    strSQL &= "AND TO_DATE(t2.YOYAKU_DT, 'YYYY/MM/DD') <= TO_DATE('" & .PropStrShiyoNenTo & "/" & .PropStrShiyoTsukiTo & "/01', 'YYYY/MM/DD') " & vbCrLf & _
                              "+ CAST('1 months' AS INTERVAL) - CAST('1 days' AS INTERVAL)"
                End If

                '利用者名が入力されている場合
                If .PropStrRiyoNm <> String.Empty Then
                    strSQL &= "AND t1.RIYO_NM LIKE '%' || '" & .PropStrRiyoNm & "' || '%'" & vbCrLf
                End If

                '利用者名カナが入力されている場合
                If .PropStrRiyoNmKana <> String.Empty Then
                    strSQL &= "AND t1.RIYO_KANA LIKE '%' || '" & .PropStrRiyoNmKana & "' || '%'"
                End If

                'ORDER BY句追加
                strSQL &= "ORDER BY t2.YOYAKU_NO, t2.YOYAKU_DT"

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

#Region "請求時の調整額取得"

    ''' <summary>
    ''' 請求時の調整額取得
    ''' </summary>
    ''' <param name="Adapter">アダプタ</param>
    ''' <param name="Cn">コネクション</param>
    ''' <param name="dataEXTZ0103">データクラス</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>請求時の調整額を取得するSQLの作成
    ''' <para>作成情報：2015/09/16 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function setSeikyuChosei(ByRef Adapter As NpgsqlDataAdapter, _
                                    ByRef Cn As NpgsqlConnection, _
                                    ByVal dataEXTZ0103 As DataEXTZ0103) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = String.Empty

        Try

            With dataEXTZ0103
                'SQL文(SELECT)
                'データアダプタに、SQLのSELECT文を設定
                If .PropStrShisetsuKbn = SHISETU_KBN_THEATER Then
                    'シアターの場合
                    strSQL = strSelectSeikyuChoseiTheatreSQL
                Else
                    'スタジオの場合
                    strSQL = strSelectSeikyuChoseiStudioSQL
                End If

                '使用年(From）・使用月(From）が共に入力されている場合
                If .PropStrShiyoNenFrom <> String.Empty AndAlso .PropStrShiyoTsukiFrom <> String.Empty Then
                    strSQL &= "AND TO_DATE(t2.YOYAKU_DT, 'YYYY/MM/DD') >= TO_DATE('" & .PropStrShiyoNenFrom & "/" & .PropStrShiyoTsukiFrom & "/01', 'YYYY/MM/DD') " & vbCrLf
                End If

                '使用年(To）・使用月(To）が共に入力されている場合
                If .PropStrShiyoNenTo <> String.Empty AndAlso .PropStrShiyoTsukiTo <> String.Empty Then
                    strSQL &= "AND TO_DATE(t2.YOYAKU_DT, 'YYYY/MM/DD') <= TO_DATE('" & .PropStrShiyoNenTo & "/" & .PropStrShiyoTsukiTo & "/01', 'YYYY/MM/DD') " & vbCrLf & _
                              "+ CAST('1 months' AS INTERVAL) - CAST('1 days' AS INTERVAL)"
                End If

                '利用者名が入力されている場合
                If .PropStrRiyoNm <> String.Empty Then
                    strSQL &= "AND t1.RIYO_NM LIKE '%' || '" & .PropStrRiyoNm & "' || '%'" & vbCrLf
                End If

                '利用者名カナが入力されている場合
                If .PropStrRiyoNmKana <> String.Empty Then
                    strSQL &= "AND t1.RIYO_KANA LIKE '%' || '" & .PropStrRiyoNmKana & "' || '%'"
                End If

                'ORDER BY句追加
                strSQL &= "ORDER BY t2.YOYAKU_NO, t2.YOYAKU_DT"

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
