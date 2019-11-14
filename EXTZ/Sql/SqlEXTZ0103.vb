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

    ' --- 2019/08/15 日別売上集計対応 Start E.Okuda@Compass ---
    ' 2019/11/07 FUTAI_KIN → FUTAI_SHOKEIに変更
    Private strSelectDayUriageTheatreSQL As String =
                "    SELECT" & vbCrLf &
                "        t1.YOYAKU_NO," & vbCrLf &
                "        t2.YOYAKU_DT," & vbCrLf &
                "        t1.SAIJI_NM," & vbCrLf &
                "        t1.RIYO_NM," & vbCrLf &
                "        CASE t1.KASHI_KIND" & vbCrLf &
                "            WHEN '0' THEN '未定'" & vbCrLf &
                "            WHEN '1' THEN '一般貸出'" & vbCrLf &
                "            WHEN '2' THEN '社内利用'" & vbCrLf &
                "            WHEN '3' THEN '特例'" & vbCrLf &
                "        END KASHI_KIND," & vbCrLf &
                "        CASE t1.RIYO_TYPE" & vbCrLf &
                "            WHEN '0' THEN '未定'" & vbCrLf &
                "            WHEN '1' THEN '着席'" & vbCrLf &
                "            WHEN '2' THEN 'スタンディング'" & vbCrLf &
                "            WHEN '3' THEN '変則'" & vbCrLf &
                "            WHEN '4' THEN '催事'" & vbCrLf &
                "        END RIYO_TYPE," & vbCrLf &
                "    CASE t1.SAIJI_BUNRUI" & vbCrLf &
                                "        WHEN '0' THEN '未定'" & vbCrLf &
                "        WHEN '1' THEN '音楽'" & vbCrLf &
                "        WHEN '2' THEN '演劇'" & vbCrLf &
                "        WHEN '3' THEN '演芸'" & vbCrLf &
                "        WHEN '4' THEN 'ビジネス'" & vbCrLf &
                "        WHEN '5' THEN '試写会・映画'" & vbCrLf &
                "        WHEN '6' THEN 'その他'" & vbCrLf &
                "    END SAIJI_BUNRUI," & vbCrLf &
                "    COALESCE(t2.RIYO_KIN,'0') AS RIYO_KIN," & vbCrLf &
                "    COALESCE(t3.FUTAI_SHOKEI,'0') AS FUTAI_SHOKEI," & vbCrLf &
                "    COALESCE(t3.FUTAI_CHOSEI,'0') AS FUTAI_CHOSEI," & vbCrLf &
                "    COALESCE(t4a.DEPOSIT_AMOUNT,'0') AS DRINK_GENKIN," & vbCrLf &
                "    COALESCE(t4b.DEPOSIT_AMOUNT,'0') AS ONE_DRINK," & vbCrLf &
                "    COALESCE(t4c.DEPOSIT_AMOUNT,'0') AS COIN_LOCKER," & vbCrLf &
                "    COALESCE(t4d.DEPOSIT_AMOUNT,'0') AS ZATSU_SHUNYU," & vbCrLf &
                "    COALESCE(t4e.DEPOSIT_AMOUNT,'0') AS AKI1," & vbCrLf &
                "    COALESCE(t4f.DEPOSIT_AMOUNT,'0') AS AKI2," & vbCrLf &
                "    COALESCE(t4g.DEPOSIT_AMOUNT,'0') AS AKI3," & vbCrLf &
                "    COALESCE(t4h.DEPOSIT_AMOUNT,'0') AS AKI4," & vbCrLf &
                "    COALESCE(t5a.DEPOSIT_AMOUNT,'0') AS SUICA_GENKIN," & vbCrLf &
                "    COALESCE(t5b.DEPOSIT_AMOUNT,'0') AS SUICA_COIN_LOCKER," & vbCrLf &
                "    COALESCE(t5c.DEPOSIT_AMOUNT,'0') AS SONOTA " & vbCrLf &
                "FROM" & vbCrLf &
                "    YOYAKU_TBL t1" & vbCrLf &
                "    LEFT JOIN" & vbCrLf &
                                "        YDT_TBL t2" & vbCrLf &
                "    ON  t1.YOYAKU_NO = t2.YOYAKU_NO" & vbCrLf &
                "    LEFT JOIN" & vbCrLf &
                "       (" & vbCrLf &
                "            SELECT" & vbCrLf &
                "                YOYAKU_NO," & vbCrLf &
                "                YOYAKU_DT," & vbCrLf &
                "                SUM(FUTAI_SHOKEI) AS FUTAI_SHOKEI," & vbCrLf &
                "                SUM(FUTAI_CHOSEI) AS FUTAI_CHOSEI" & vbCrLf &
                "            FROM" & vbCrLf &
                "                FRIYO_MEISAI_TBL" & vbCrLf &
                                "            GROUP BY" & vbCrLf &
                "                YOYAKU_NO," & vbCrLf &
                "                YOYAKU_DT" & vbCrLf &
                "        ) t3" & vbCrLf &
                "    ON  t1.YOYAKU_NO = t3.YOYAKU_NO" & vbCrLf &
                "        AND t2.YOYAKU_DT = t3.YOYAKU_DT" & vbCrLf &
                "    LEFT JOIN" & vbCrLf &
                "        (" & vbCrLf &
                "            SELECT" & vbCrLf &
                "                YOYAKU_NO," & vbCrLf &
                "                DEPOSIT_DT," & vbCrLf &
                "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf &
                "            FROM" & vbCrLf &
                "                ALSOK_DEPOSIT_TBL" & vbCrLf &
                "            WHERE" & vbCrLf &
                "                DEPOSIT_KBN = '1'" & vbCrLf &
                "            AND REGISTER_CD= '001'" & vbCrLf &
                "            GROUP BY" & vbCrLf &
                "                YOYAKU_NO," & vbCrLf &
                "                DEPOSIT_DT" & vbCrLf &
                "        ) t4a" & vbCrLf &
                "    ON  t1.YOYAKU_NO = t4a.YOYAKU_NO" & vbCrLf &
                "        AND t2.YOYAKU_DT = t4a.DEPOSIT_DT" & vbCrLf &
                "    LEFT JOIN" & vbCrLf &
                "        (" & vbCrLf &
                "            SELECT" & vbCrLf &
                "                YOYAKU_NO," & vbCrLf &
                "                DEPOSIT_DT," & vbCrLf &
                "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf &
                "            FROM" & vbCrLf &
                "                ALSOK_DEPOSIT_TBL" & vbCrLf &
                "            WHERE" & vbCrLf &
                "                DEPOSIT_KBN = '1'" & vbCrLf &
                "            AND REGISTER_CD= '002'" & vbCrLf &
                "            GROUP BY" & vbCrLf &
                "                YOYAKU_NO," & vbCrLf &
                "                DEPOSIT_DT" & vbCrLf &
                "        ) t4b" & vbCrLf &
                "    ON  t1.YOYAKU_NO = t4b.YOYAKU_NO" & vbCrLf &
                "        AND t2.YOYAKU_DT = t4b.DEPOSIT_DT" & vbCrLf &
                "    LEFT JOIN" & vbCrLf &
                "        (" & vbCrLf &
                "            SELECT" & vbCrLf &
                "                YOYAKU_NO," & vbCrLf &
                "                DEPOSIT_DT," & vbCrLf &
                "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf &
                "            FROM" & vbCrLf &
                "                ALSOK_DEPOSIT_TBL" & vbCrLf &
                "            WHERE" & vbCrLf &
                "                DEPOSIT_KBN = '1'" & vbCrLf &
                "            AND REGISTER_CD= '003'" & vbCrLf &
                "            GROUP BY" & vbCrLf &
                "                YOYAKU_NO," & vbCrLf &
                "                DEPOSIT_DT" & vbCrLf &
                "        ) t4c" & vbCrLf &
                "    ON  t1.YOYAKU_NO = t4c.YOYAKU_NO" & vbCrLf &
                "        AND t2.YOYAKU_DT = t4c.DEPOSIT_DT" & vbCrLf &
                "    LEFT JOIN" & vbCrLf &
                "        (" & vbCrLf &
                "            SELECT" & vbCrLf &
                "                YOYAKU_NO," & vbCrLf &
                "                DEPOSIT_DT," & vbCrLf &
                "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf &
                "            FROM" & vbCrLf &
                "                ALSOK_DEPOSIT_TBL" & vbCrLf &
                "            WHERE" & vbCrLf &
                "                DEPOSIT_KBN = '1'" & vbCrLf &
                "            AND REGISTER_CD= '004'" & vbCrLf &
                "            GROUP BY" & vbCrLf &
                "                YOYAKU_NO," & vbCrLf &
                "                DEPOSIT_DT" & vbCrLf &
                "        ) t4d" & vbCrLf &
                "    ON  t1.YOYAKU_NO = t4d.YOYAKU_NO" & vbCrLf &
                "        AND t2.YOYAKU_DT = t4d.DEPOSIT_DT" & vbCrLf &
                "    LEFT JOIN" & vbCrLf &
                "        (" & vbCrLf &
                "            SELECT" & vbCrLf &
                "                YOYAKU_NO," & vbCrLf &
                "                DEPOSIT_DT," & vbCrLf &
                "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf &
                "            FROM" & vbCrLf &
                "                ALSOK_DEPOSIT_TBL" & vbCrLf &
                "            WHERE" & vbCrLf &
                "                DEPOSIT_KBN = '1'" & vbCrLf &
                "            AND REGISTER_CD= '006'" & vbCrLf &
                "            GROUP BY" & vbCrLf &
                "                YOYAKU_NO," & vbCrLf &
                "                DEPOSIT_DT" & vbCrLf &
                "        ) t4e" & vbCrLf &
                "    ON  t1.YOYAKU_NO = t4e.YOYAKU_NO" & vbCrLf &
                "        AND t2.YOYAKU_DT = t4e.DEPOSIT_DT" & vbCrLf &
                "    LEFT JOIN" & vbCrLf &
                "        (" & vbCrLf &
                "            SELECT" & vbCrLf &
                "                YOYAKU_NO," & vbCrLf &
                "                DEPOSIT_DT," & vbCrLf &
                "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf &
                "            FROM" & vbCrLf &
                "                ALSOK_DEPOSIT_TBL" & vbCrLf &
                "            WHERE" & vbCrLf &
                "                DEPOSIT_KBN = '1'" & vbCrLf &
                "            AND REGISTER_CD= '007'" & vbCrLf &
                "            GROUP BY" & vbCrLf &
                "                YOYAKU_NO," & vbCrLf &
                "                DEPOSIT_DT" & vbCrLf &
                "        ) t4f" & vbCrLf &
                "    ON  t1.YOYAKU_NO = t4f.YOYAKU_NO" & vbCrLf &
                "        AND t2.YOYAKU_DT = t4f.DEPOSIT_DT" & vbCrLf &
                "    LEFT JOIN" & vbCrLf &
                "        (" & vbCrLf &
                "            SELECT" & vbCrLf &
                "                YOYAKU_NO," & vbCrLf &
                "                DEPOSIT_DT," & vbCrLf &
                "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf &
                "            FROM" & vbCrLf &
                "                ALSOK_DEPOSIT_TBL" & vbCrLf &
                "            WHERE" & vbCrLf &
                "                DEPOSIT_KBN = '1'" & vbCrLf &
                "            AND REGISTER_CD= '008'" & vbCrLf &
                "            GROUP BY" & vbCrLf &
                "                YOYAKU_NO," & vbCrLf &
                "                DEPOSIT_DT" & vbCrLf &
                "       ) t4g" & vbCrLf &
                "    ON  t1.YOYAKU_NO = t4g.YOYAKU_NO" & vbCrLf &
                "        AND t2.YOYAKU_DT = t4g.DEPOSIT_DT" & vbCrLf &
                "    LEFT JOIN" & vbCrLf &
                "        (" & vbCrLf &
                "            SELECT" & vbCrLf &
                "                YOYAKU_NO," & vbCrLf &
                "                DEPOSIT_DT," & vbCrLf &
                "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf &
                "            FROM" & vbCrLf &
                "                ALSOK_DEPOSIT_TBL" & vbCrLf &
                "            WHERE" & vbCrLf &
                "                DEPOSIT_KBN = '1'" & vbCrLf &
                "            AND REGISTER_CD= '009'" & vbCrLf &
                "            GROUP BY" & vbCrLf &
                "                YOYAKU_NO," & vbCrLf &
                "                DEPOSIT_DT" & vbCrLf &
                "        ) t4h" & vbCrLf &
                "    ON  t1.YOYAKU_NO = t4h.YOYAKU_NO" & vbCrLf &
                "        AND t2.YOYAKU_DT = t4h.DEPOSIT_DT" & vbCrLf &
                "    LEFT JOIN" & vbCrLf &
                "        (" & vbCrLf &
                "            SELECT" & vbCrLf &
                "                YOYAKU_NO," & vbCrLf &
                "                DEPOSIT_DT," & vbCrLf &
                "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf &
                "            FROM" & vbCrLf &
                "                ALSOK_DEPOSIT_TBL" & vbCrLf &
                "            WHERE" & vbCrLf &
                "                DEPOSIT_KBN = '2'" & vbCrLf &
                "            AND REGISTER_CD= '900'" & vbCrLf &
                "            GROUP BY" & vbCrLf &
                "                YOYAKU_NO," & vbCrLf &
                "                DEPOSIT_DT" & vbCrLf &
                "        ) t5a" & vbCrLf &
                "    ON  t1.YOYAKU_NO = t5a.YOYAKU_NO" & vbCrLf &
                "        AND t2.YOYAKU_DT = t5a.DEPOSIT_DT" & vbCrLf &
                "    LEFT JOIN" & vbCrLf &
                "        (" & vbCrLf &
                "            SELECT" & vbCrLf &
                "                YOYAKU_NO," & vbCrLf &
                "                DEPOSIT_DT," & vbCrLf &
                "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf &
                "            FROM" & vbCrLf &
                "                ALSOK_DEPOSIT_TBL" & vbCrLf &
                "            WHERE" & vbCrLf &
                "                DEPOSIT_KBN = '2'" & vbCrLf &
                "            AND REGISTER_CD= '901'" & vbCrLf &
                "            GROUP BY" & vbCrLf &
                "                YOYAKU_NO," & vbCrLf &
                "                DEPOSIT_DT" & vbCrLf &
                "        ) t5b" & vbCrLf &
                "    ON  t1.YOYAKU_NO = t5b.YOYAKU_NO" & vbCrLf &
                "        AND t2.YOYAKU_DT = t5b.DEPOSIT_DT" & vbCrLf &
                "    LEFT JOIN" & vbCrLf &
                "        (" & vbCrLf &
                "            SELECT" & vbCrLf &
                "                YOYAKU_NO," & vbCrLf &
                "                DEPOSIT_DT," & vbCrLf &
                "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf &
                "            FROM" & vbCrLf &
                "                ALSOK_DEPOSIT_TBL" & vbCrLf &
                "            WHERE" & vbCrLf &
                "                DEPOSIT_KBN = '2'" & vbCrLf &
                "            AND REGISTER_CD BETWEEN '902' AND '950'" & vbCrLf &
                "            GROUP BY" & vbCrLf &
                "                YOYAKU_NO," & vbCrLf &
                "                DEPOSIT_DT" & vbCrLf &
                "        ) t5c" & vbCrLf &
                "    ON  t1.YOYAKU_NO = t5c.YOYAKU_NO" & vbCrLf &
                "        AND t2.YOYAKU_DT = t5c.DEPOSIT_DT" & vbCrLf &
                "    WHERE " & vbCrLf &
                "        t2.SHISETU_KBN = '1'" & vbCrLf

    'Private strSelectDayUriageTheatreSQL As String = _
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
    '            "        WHEN '1' THEN '着席'" & vbCrLf & _
    '            "        WHEN '2' THEN 'スタンディング'" & vbCrLf & _
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
    ' 2017.04.17 e.watanabe UPD END↑ シアター利用形状の修正（1:着席、2:スタンディング）
    ' --- 2019/08/15 日別売上集計対応 End E.Okuda@Compass ---

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

    ' 2019/11/07 FUTAI_KIN → FUTAI_SHOKEIに変更
    Private strSelectDayUriageStudioSQL As String =
                "SELECT" & vbCrLf &
                "    t1.YOYAKU_NO," & vbCrLf &
                "    t2.YOYAKU_DT," & vbCrLf &
                "    t1.SHUTSUEN_NM," & vbCrLf &
                "    t1.RIYO_NM," & vbCrLf &
                "    CASE t1.KASHI_KIND" & vbCrLf &
                "        WHEN '0' THEN '未定'" & vbCrLf &
                "        WHEN '1' THEN '一般貸出'" & vbCrLf &
                "        WHEN '2' THEN '社内利用'" & vbCrLf &
                "    END KASHI_KIND," & vbCrLf &
                "    CASE t1.STUDIO_KBN" & vbCrLf &
                "        WHEN '1' THEN '201st'" & vbCrLf &
                "        WHEN '2' THEN '202st'" & vbCrLf &
                "        WHEN '3' THEN 'house lock'" & vbCrLf &
                "    END STUDIO_KBN," & vbCrLf &
                "    COALESCE(t2.RIYO_KIN,'0') AS RIYO_KIN," & vbCrLf &
                "    COALESCE(t3.FUTAI_SHOKEI,'0') AS FUTAI_SHOKEI," & vbCrLf &
                "    COALESCE(t3.FUTAI_CHOSEI,'0') AS FUTAI_CHOSEI," & vbCrLf &
                "    COALESCE(t4a.DEPOSIT_AMOUNT,'0') AS DRINK_GENKIN," & vbCrLf &
                "    COALESCE(t4b.DEPOSIT_AMOUNT,'0') AS ONE_DRINK," & vbCrLf &
                "    COALESCE(t4c.DEPOSIT_AMOUNT,'0') AS COIN_LOCKER," & vbCrLf &
                "    COALESCE(t4d.DEPOSIT_AMOUNT,'0') AS ZATSU_SHUNYU," & vbCrLf &
                "    COALESCE(t4e.DEPOSIT_AMOUNT,'0') AS AKI1," & vbCrLf &
                "    COALESCE(t4f.DEPOSIT_AMOUNT,'0') AS AKI2," & vbCrLf &
                "    COALESCE(t4g.DEPOSIT_AMOUNT,'0') AS AKI3," & vbCrLf &
                "    COALESCE(t4h.DEPOSIT_AMOUNT,'0') AS AKI4," & vbCrLf &
                "    COALESCE(t5a.DEPOSIT_AMOUNT,'0') AS SUICA_GENKIN," & vbCrLf &
                "    COALESCE(t5b.DEPOSIT_AMOUNT,'0') AS SUICA_COIN_LOCKER," & vbCrLf &
                "    COALESCE(t5c.DEPOSIT_AMOUNT,'0') AS SONOTA" & vbCrLf &
                "FROM" & vbCrLf &
                "    YOYAKU_TBL t1" & vbCrLf &
                "    LEFT JOIN" & vbCrLf &
                "        YDT_TBL t2" & vbCrLf &
                "    ON  t1.YOYAKU_NO = t2.YOYAKU_NO" & vbCrLf &
                "    LEFT JOIN" & vbCrLf &
                "       (" & vbCrLf &
                "            SELECT" & vbCrLf &
                "                YOYAKU_NO," & vbCrLf &
                "                YOYAKU_DT," & vbCrLf &
                "                SUM(FUTAI_SHOKEI) AS FUTAI_SHOKEI," & vbCrLf &
                "                SUM(FUTAI_CHOSEI) AS FUTAI_CHOSEI" & vbCrLf &
                "            FROM" & vbCrLf &
                "                FRIYO_MEISAI_TBL" & vbCrLf &
                "            GROUP BY" & vbCrLf &
                "                YOYAKU_NO," & vbCrLf &
                "                YOYAKU_DT" & vbCrLf &
                "        ) t3" & vbCrLf &
                "    ON  t1.YOYAKU_NO = t3.YOYAKU_NO" & vbCrLf &
                "        AND t2.YOYAKU_DT = t3.YOYAKU_DT" & vbCrLf &
                "    LEFT JOIN" & vbCrLf &
                "        (" & vbCrLf &
                "            SELECT" & vbCrLf &
                "                YOYAKU_NO," & vbCrLf &
                "                DEPOSIT_DT," & vbCrLf &
                "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf &
                "            FROM" & vbCrLf &
                "                ALSOK_DEPOSIT_TBL" & vbCrLf &
                "            WHERE" & vbCrLf &
                "                DEPOSIT_KBN = '1'" & vbCrLf &
                "            AND REGISTER_CD= '001'" & vbCrLf &
                "            GROUP BY" & vbCrLf &
                "                YOYAKU_NO," & vbCrLf &
                "                DEPOSIT_DT" & vbCrLf &
                "        ) t4a" & vbCrLf &
                "    ON  t1.YOYAKU_NO = t4a.YOYAKU_NO" & vbCrLf &
                "        AND t2.YOYAKU_DT = t4a.DEPOSIT_DT" & vbCrLf &
                "    LEFT JOIN" & vbCrLf &
                "        (" & vbCrLf &
                "            SELECT" & vbCrLf &
                "                YOYAKU_NO," & vbCrLf &
                "                DEPOSIT_DT," & vbCrLf &
                "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf &
                "            FROM" & vbCrLf &
                "                ALSOK_DEPOSIT_TBL" & vbCrLf &
                "            WHERE" & vbCrLf &
                "                DEPOSIT_KBN = '1'" & vbCrLf &
                "            AND REGISTER_CD= '002'" & vbCrLf &
                "            GROUP BY" & vbCrLf &
                "                YOYAKU_NO," & vbCrLf &
                "                DEPOSIT_DT" & vbCrLf &
                "        ) t4b" & vbCrLf &
                "    ON  t1.YOYAKU_NO = t4b.YOYAKU_NO" & vbCrLf &
                "        AND t2.YOYAKU_DT = t4b.DEPOSIT_DT" & vbCrLf &
                "    LEFT JOIN" & vbCrLf &
                "        (" & vbCrLf &
                "            SELECT" & vbCrLf &
                "                YOYAKU_NO," & vbCrLf &
                "                DEPOSIT_DT," & vbCrLf &
                "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf &
                "            FROM" & vbCrLf &
                "                ALSOK_DEPOSIT_TBL" & vbCrLf &
                "            WHERE" & vbCrLf &
                "                DEPOSIT_KBN = '1'" & vbCrLf &
                "            AND REGISTER_CD= '003'" & vbCrLf &
                "            GROUP BY" & vbCrLf &
                "                YOYAKU_NO," & vbCrLf &
                "                DEPOSIT_DT" & vbCrLf &
                "        ) t4c" & vbCrLf &
                "    ON  t1.YOYAKU_NO = t4c.YOYAKU_NO" & vbCrLf &
                "        AND t2.YOYAKU_DT = t4c.DEPOSIT_DT" & vbCrLf &
                "    LEFT JOIN" & vbCrLf &
                "        (" & vbCrLf &
                "            SELECT" & vbCrLf &
                "                YOYAKU_NO," & vbCrLf &
                "                DEPOSIT_DT," & vbCrLf &
                "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf &
                "            FROM" & vbCrLf &
                "                ALSOK_DEPOSIT_TBL" & vbCrLf &
                "            WHERE" & vbCrLf &
                "                DEPOSIT_KBN = '1'" & vbCrLf &
                "            AND REGISTER_CD= '004'" & vbCrLf &
                "            GROUP BY" & vbCrLf &
                "                YOYAKU_NO," & vbCrLf &
                "                DEPOSIT_DT" & vbCrLf &
                "        ) t4d" & vbCrLf &
                "    ON  t1.YOYAKU_NO = t4d.YOYAKU_NO" & vbCrLf &
                "        AND t2.YOYAKU_DT = t4d.DEPOSIT_DT" & vbCrLf &
                "    LEFT JOIN" & vbCrLf &
                "        (" & vbCrLf &
                "            SELECT" & vbCrLf &
                "                YOYAKU_NO," & vbCrLf &
                "                DEPOSIT_DT," & vbCrLf &
                "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf &
                "            FROM" & vbCrLf &
                "                ALSOK_DEPOSIT_TBL" & vbCrLf &
                "            WHERE" & vbCrLf &
                "                DEPOSIT_KBN = '1'" & vbCrLf &
                "            AND REGISTER_CD= '006'" & vbCrLf &
                "            GROUP BY" & vbCrLf &
                "                YOYAKU_NO," & vbCrLf &
                "                DEPOSIT_DT" & vbCrLf &
                "        ) t4e" & vbCrLf &
                "    ON  t1.YOYAKU_NO = t4e.YOYAKU_NO" & vbCrLf &
                "        AND t2.YOYAKU_DT = t4e.DEPOSIT_DT" & vbCrLf &
                "    LEFT JOIN" & vbCrLf &
                "        (" & vbCrLf &
                "            SELECT" & vbCrLf &
                "                YOYAKU_NO," & vbCrLf &
                "                DEPOSIT_DT," & vbCrLf &
                "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf &
                "            FROM" & vbCrLf &
                "                ALSOK_DEPOSIT_TBL" & vbCrLf &
                "            WHERE" & vbCrLf &
                "                DEPOSIT_KBN = '1'" & vbCrLf &
                "            AND REGISTER_CD= '007'" & vbCrLf &
                "            GROUP BY" & vbCrLf &
                "                YOYAKU_NO," & vbCrLf &
                "                DEPOSIT_DT" & vbCrLf &
                "        ) t4f" & vbCrLf &
                "    ON  t1.YOYAKU_NO = t4f.YOYAKU_NO" & vbCrLf &
                "        AND t2.YOYAKU_DT = t4f.DEPOSIT_DT" & vbCrLf &
                "    LEFT JOIN" & vbCrLf &
                "        (" & vbCrLf &
                "            SELECT" & vbCrLf &
                "                YOYAKU_NO," & vbCrLf &
                "                DEPOSIT_DT," & vbCrLf &
                "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf &
                "            FROM" & vbCrLf &
                "                ALSOK_DEPOSIT_TBL" & vbCrLf &
                "            WHERE" & vbCrLf &
                "                DEPOSIT_KBN = '1'" & vbCrLf &
                "            AND REGISTER_CD= '008'" & vbCrLf &
                "            GROUP BY" & vbCrLf &
                "                YOYAKU_NO," & vbCrLf &
                "                DEPOSIT_DT" & vbCrLf &
                "        ) t4g" & vbCrLf &
                "    ON  t1.YOYAKU_NO = t4g.YOYAKU_NO" & vbCrLf &
                "        AND t2.YOYAKU_DT = t4g.DEPOSIT_DT" & vbCrLf &
                "    LEFT JOIN" & vbCrLf &
                "        (" & vbCrLf &
                "            SELECT" & vbCrLf &
                "                YOYAKU_NO," & vbCrLf &
                "                DEPOSIT_DT," & vbCrLf &
                "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf &
                "            FROM" & vbCrLf &
                "                ALSOK_DEPOSIT_TBL" & vbCrLf &
                "            WHERE" & vbCrLf &
                "                DEPOSIT_KBN = '1'" & vbCrLf &
                "            AND REGISTER_CD= '009'" & vbCrLf &
                "            GROUP BY" & vbCrLf &
                "                YOYAKU_NO," & vbCrLf &
                "                DEPOSIT_DT" & vbCrLf &
                "        ) t4h" & vbCrLf &
                "    ON  t1.YOYAKU_NO = t4h.YOYAKU_NO" & vbCrLf &
                "        AND t2.YOYAKU_DT = t4h.DEPOSIT_DT" & vbCrLf &
                "    LEFT JOIN" & vbCrLf &
                "        (" & vbCrLf &
                "            SELECT" & vbCrLf &
                "                YOYAKU_NO," & vbCrLf &
                "                DEPOSIT_DT," & vbCrLf &
                "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf &
                "            FROM" & vbCrLf &
                "                ALSOK_DEPOSIT_TBL" & vbCrLf &
                "            WHERE" & vbCrLf &
                "                DEPOSIT_KBN = '2'" & vbCrLf &
                "            AND REGISTER_CD= '900'" & vbCrLf &
                "            GROUP BY" & vbCrLf &
                "                YOYAKU_NO," & vbCrLf &
                "                DEPOSIT_DT" & vbCrLf &
                "        ) t5a" & vbCrLf &
                "    ON  t1.YOYAKU_NO = t5a.YOYAKU_NO" & vbCrLf &
                "        AND t2.YOYAKU_DT = t5a.DEPOSIT_DT" & vbCrLf &
                "    LEFT JOIN" & vbCrLf &
                "        (" & vbCrLf &
                "            SELECT" & vbCrLf &
                "                YOYAKU_NO," & vbCrLf &
                "                DEPOSIT_DT," & vbCrLf &
                "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf &
                "            FROM" & vbCrLf &
                "                ALSOK_DEPOSIT_TBL" & vbCrLf &
                "            WHERE" & vbCrLf &
                "                DEPOSIT_KBN = '2'" & vbCrLf &
                "            AND REGISTER_CD= '901'" & vbCrLf &
                "            GROUP BY" & vbCrLf &
                "                YOYAKU_NO," & vbCrLf &
                "                DEPOSIT_DT" & vbCrLf &
                "        ) t5b" & vbCrLf &
                "    ON  t1.YOYAKU_NO = t5b.YOYAKU_NO" & vbCrLf &
                "        AND t2.YOYAKU_DT = t5b.DEPOSIT_DT" & vbCrLf &
                "    LEFT JOIN" & vbCrLf &
                "        (" & vbCrLf &
                "            SELECT" & vbCrLf &
                "                YOYAKU_NO," & vbCrLf &
                "                DEPOSIT_DT," & vbCrLf &
                "                SUM(DEPOSIT_AMOUNT) AS DEPOSIT_AMOUNT" & vbCrLf &
                "            FROM" & vbCrLf &
                "                ALSOK_DEPOSIT_TBL" & vbCrLf &
                "            WHERE" & vbCrLf &
                "                DEPOSIT_KBN = '2'" & vbCrLf &
                "            AND REGISTER_CD BETWEEN '902' AND '950'" & vbCrLf &
                "            GROUP BY" & vbCrLf &
                "                YOYAKU_NO," & vbCrLf &
                "                DEPOSIT_DT" & vbCrLf &
                "        ) t5c" & vbCrLf &
                "    ON  t1.YOYAKU_NO = t5c.YOYAKU_NO" & vbCrLf &
                "        AND t2.YOYAKU_DT = t5c.DEPOSIT_DT" & vbCrLf &
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

                ' --- 2019/08/23 日別売上一覧機能追加対応 Start E.Okuda@Compass ---
                '利用者名カナが入力されている場合
                'If .PropStrRiyoNmKana <> String.Empty Then
                '    strSQL &= "AND t1.RIYO_KANA LIKE '%' || '" & .PropStrRiyoNmKana & "' || '%'"
                'End If
                If .PropStrRiyoNmKana <> String.Empty Then
                    strSQL &= "AND t1.RIYO_KANA LIKE '%' || '" & .PropStrRiyoNmKana & "' || '%'" & vbCrLf
                End If

                ' 利用料／付帯設備利用料ゼロ除外チェックが入っている場合
                If .PropBlnExcludeZero Then
                    strSQL &= "AND (RIYO_KIN <> 0 OR FUTAI_SHOKEI <> 0)" & vbCrLf
                End If

                'ORDER BY句追加
                strSQL &= " ORDER BY t2.YOYAKU_DT, t1.SAIJI_NM"
                '                strSQL &= "ORDER BY t2.YOYAKU_NO, t2.YOYAKU_DT"
                ' --- 2019/08/23 日別売上一覧機能追加対応 End E.Okuda@Compass ---

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
