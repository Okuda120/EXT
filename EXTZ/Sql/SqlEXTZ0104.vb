Imports Common
Imports CommonEXT
Imports Npgsql

''' <summary>
''' 利用状況一覧で使用する情報取得SQL作成
''' </summary>
''' <remarks>利用状況一覧で使用する情報取得のSQLを作成する
''' <para>作成情報：2015/09/24 h.endo
''' <p>改訂情報:</p>
''' </para></remarks>
Public Class SqlEXTZ0104

    '*****基本SQL文*****
#Region "SQL文(利用状況一覧(シアター)取得)"

    'SQL文(利用状況一覧(シアター)取得)
    ' 2017.07.31 e.watanabe UPD START↓ シアター利用形状の修正（1:使用日数<着席>、2:使用日数<スタンディング>）
    '    Private strSelectRiyoJokyoTheatreSQL As String = _
    '"    SELECT" & vbCrLf & _
    '"        EXTRACT(DAY FROM (CAST(:HYOJI_NEN ||:HYOJI_TSUKI || '01' as date) + cast('1 Month' as interval) - cast('1 day' as interval)))" & vbCrLf & _
    '"         - t1.COUNT - t2.COUNT AS RIYO_KANO_NISSU," & vbCrLf & _
    '"        t1.COUNT AS KYUKAN_NISSU," & vbCrLf & _
    '"        t2.COUNT AS MAINTE_NISSU," & vbCrLf & _
    '"        t4a.COUNT AS RIYO_NISSU_CHAKUSEKI," & vbCrLf & _
    '"        t4b.COUNT AS RIYO_NISSU_STANDING," & vbCrLf & _
    '"        t4c.COUNT AS RIYO_NISSU_MIX," & vbCrLf & _
    '"        t4d.COUNT AS RIYO_NISSU_SAIJI," & vbCrLf & _
    '"        t5a.COUNT AS RIYO_NISSU_MUSIC," & vbCrLf & _
    '"        t5b.COUNT AS RIYO_NISSU_ENGEKI," & vbCrLf & _
    '"        t5c.COUNT AS RIYO_NISSU_ENGEI," & vbCrLf & _
    '"        t5d.COUNT AS RIYO_NISSU_BUSINESS," & vbCrLf & _
    '"        t5e.COUNT AS RIYO_NISSU_MOVIE," & vbCrLf & _
    '"        t5f.COUNT AS RIYO_NISSU_SONOTA" & vbCrLf & _
    '"    FROM" & vbCrLf & _
    '"        (" & vbCrLf & _
    '"          SELECT" & vbCrLf & _
    '"           COUNT(TMP.HOLMENT_DT) AS COUNT" & vbCrLf & _
    '"           FROM" & vbCrLf & _
    '"           (" & vbCrLf & _
    '"            SELECT" & vbCrLf & _
    '"                HOLMENT_DT" & vbCrLf & _
    '"            FROM" & vbCrLf & _
    '"                HOLMENT_MST" & vbCrLf & _
    '"            WHERE" & vbCrLf & _
    '"            SHISETU_KBN = '1'" & vbCrLf & _
    '"            AND HOLMENT_KBN = '1'" & vbCrLf & _
    '"            AND TO_DATE(HOLMENT_DT, 'YYYY/MM/DD') BETWEEN CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' AS DATE)    " & vbCrLf & _
    '"                AND CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' as date) + cast('1 Month' as interval) - cast('1 day' as interval)" & vbCrLf & _
    '"            GROUP BY HOLMENT_DT" & vbCrLf & _
    '"           ) TMP" & vbCrLf & _
    '"        ) t1" & vbCrLf & _
    '"    CROSS JOIN" & vbCrLf & _
    '"        (" & vbCrLf & _
    '"          SELECT" & vbCrLf & _
    '"           COUNT(TMP.HOLMENT_DT) AS COUNT" & vbCrLf & _
    '"           FROM" & vbCrLf & _
    '"           (" & vbCrLf & _
    '"            SELECT" & vbCrLf & _
    '"                HOLMENT_DT" & vbCrLf & _
    '"            FROM" & vbCrLf & _
    '"                HOLMENT_MST" & vbCrLf & _
    '"            WHERE" & vbCrLf & _
    '"            SHISETU_KBN = '1'" & vbCrLf & _
    '"            AND HOLMENT_KBN = '2'" & vbCrLf & _
    '"            AND TO_DATE(HOLMENT_DT, 'YYYY/MM/DD') BETWEEN CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' AS DATE)    " & vbCrLf & _
    '"                AND CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' as date) + cast('1 Month' as interval) - cast('1 day' as interval)" & vbCrLf & _
    '"            GROUP BY HOLMENT_DT" & vbCrLf & _
    '"           ) TMP" & vbCrLf & _
    '"        ) t2" & vbCrLf & _
    '"    CROSS JOIN" & vbCrLf & _
    '"        (" & vbCrLf & _
    '"          SELECT" & vbCrLf & _
    '"           COUNT(TMP.YOYAKU_DT) AS COUNT" & vbCrLf & _
    '"           FROM" & vbCrLf & _
    '"           (" & vbCrLf & _
    '"            SELECT" & vbCrLf & _
    '"                t4a2.YOYAKU_DT" & vbCrLf & _
    '"            FROM" & vbCrLf & _
    '"                YOYAKU_TBL t4a1" & vbCrLf & _
    '"                LEFT JOIN" & vbCrLf & _
    '"                    YDT_TBL t4a2" & vbCrLf & _
    '"                ON  t4a1.YOYAKU_NO = t4a2.YOYAKU_NO" & vbCrLf & _
    '"            WHERE" & vbCrLf & _
    '"                t4a1.YOYAKU_STS = '4'" & vbCrLf & _
    '"            AND t4a1.SHISETU_KBN = '1'" & vbCrLf & _
    '"            AND t4a1.RIYO_TYPE = '2'" & vbCrLf & _
    '"            AND TO_DATE(t4a2.YOYAKU_DT, 'YYYY/MM/DD') BETWEEN CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' AS DATE)    " & vbCrLf & _
    '"                AND CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' as date) + cast('1 Month' as interval) - cast('1 day' as interval)" & vbCrLf & _
    '"            GROUP BY" & vbCrLf & _
    '"                t4a2.YOYAKU_DT" & vbCrLf & _
    '"           ) TMP" & vbCrLf & _
    '"        ) t4a" & vbCrLf & _
    '"    CROSS JOIN" & vbCrLf & _
    '"        (" & vbCrLf & _
    '"          SELECT" & vbCrLf & _
    '"           COUNT(TMP.YOYAKU_DT) AS COUNT" & vbCrLf & _
    '"           FROM" & vbCrLf & _
    '"           (" & vbCrLf & _
    '"                SELECT" & vbCrLf & _
    '"                    t4b2.YOYAKU_DT" & vbCrLf & _
    '"                FROM" & vbCrLf & _
    '"                    YOYAKU_TBL t4b1" & vbCrLf & _
    '"                    LEFT JOIN" & vbCrLf & _
    '"                        YDT_TBL t4b2" & vbCrLf & _
    '"                    ON  t4b1.YOYAKU_NO = t4b2.YOYAKU_NO" & vbCrLf & _
    '"                WHERE" & vbCrLf & _
    '"                    t4b1.YOYAKU_STS = '4'" & vbCrLf & _
    '"                AND t4b1.SHISETU_KBN = '1'" & vbCrLf & _
    '"                AND t4b1.RIYO_TYPE = '1'" & vbCrLf & _
    '"                AND TO_DATE(t4b2.YOYAKU_DT, 'YYYY/MM/DD') BETWEEN CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' AS DATE)    " & vbCrLf & _
    '"                AND CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' as date) + cast('1 Month' as interval) - cast('1 day' as interval)" & vbCrLf & _
    '"                GROUP BY" & vbCrLf & _
    '"                    t4b2.YOYAKU_DT" & vbCrLf & _
    '"           ) TMP" & vbCrLf & _
    '"        ) t4b" & vbCrLf & _
    '"    CROSS JOIN" & vbCrLf & _
    '"        (" & vbCrLf & _
    '"          SELECT" & vbCrLf & _
    '"           COUNT(TMP.YOYAKU_DT) AS COUNT" & vbCrLf & _
    '"           FROM" & vbCrLf & _
    '"           (" & vbCrLf & _
    '"                SELECT" & vbCrLf & _
    '"                    t4c2.YOYAKU_DT" & vbCrLf & _
    '"                FROM" & vbCrLf & _
    '"                    YOYAKU_TBL t4c1" & vbCrLf & _
    '"                    LEFT JOIN" & vbCrLf & _
    '"                        YDT_TBL t4c2" & vbCrLf & _
    '"                    ON  t4c1.YOYAKU_NO = t4c2.YOYAKU_NO" & vbCrLf & _
    '"                WHERE" & vbCrLf & _
    '"                    t4c1.YOYAKU_STS = '4'" & vbCrLf & _
    '"                AND t4c1.SHISETU_KBN = '1'" & vbCrLf & _
    '"                AND t4c1.RIYO_TYPE = '3'" & vbCrLf & _
    '"                AND TO_DATE(t4c2.YOYAKU_DT, 'YYYY/MM/DD') BETWEEN CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' AS DATE)    " & vbCrLf & _
    '"                AND CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' as date) + cast('1 Month' as interval) - cast('1 day' as interval)" & vbCrLf & _
    '"                GROUP BY" & vbCrLf & _
    '"                    t4c2.YOYAKU_DT" & vbCrLf & _
    '"           ) TMP" & vbCrLf & _
    '"        ) t4c" & vbCrLf & _
    '"    CROSS JOIN" & vbCrLf & _
    '"        (" & vbCrLf & _
    '"          SELECT" & vbCrLf & _
    '"           COUNT(TMP.YOYAKU_DT) AS COUNT" & vbCrLf & _
    '"           FROM" & vbCrLf & _
    '"           (" & vbCrLf & _
    '"                SELECT" & vbCrLf & _
    '"                    t4d2.YOYAKU_DT" & vbCrLf & _
    '"                FROM" & vbCrLf & _
    '"                    YOYAKU_TBL t4d1" & vbCrLf & _
    '"                    LEFT JOIN" & vbCrLf & _
    '"                        YDT_TBL t4d2" & vbCrLf & _
    '"                    ON  t4d1.YOYAKU_NO = t4d2.YOYAKU_NO" & vbCrLf & _
    '"                WHERE" & vbCrLf & _
    '"                    t4d1.YOYAKU_STS = '4'" & vbCrLf & _
    '"                AND t4d1.SHISETU_KBN = '1'" & vbCrLf & _
    '"                AND t4d1.RIYO_TYPE = '4'" & vbCrLf & _
    '"                AND TO_DATE(t4d2.YOYAKU_DT, 'YYYY/MM/DD') BETWEEN CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' AS DATE)    " & vbCrLf & _
    '"                AND CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' as date) + cast('1 Month' as interval) - cast('1 day' as interval)" & vbCrLf & _
    '"                GROUP BY" & vbCrLf & _
    '"                    t4d2.YOYAKU_DT" & vbCrLf & _
    '"           ) TMP" & vbCrLf & _
    '"        ) t4d" & vbCrLf & _
    '"    CROSS JOIN" & vbCrLf & _
    '"        (" & vbCrLf & _
    '"          SELECT" & vbCrLf & _
    '"           COUNT(TMP.YOYAKU_DT) AS COUNT" & vbCrLf & _
    '"           FROM" & vbCrLf & _
    '"           (" & vbCrLf & _
    '"                SELECT" & vbCrLf & _
    '"                    t5a2.YOYAKU_DT" & vbCrLf & _
    '"                FROM" & vbCrLf & _
    '"                    YOYAKU_TBL t5a1" & vbCrLf & _
    '"                    LEFT JOIN" & vbCrLf & _
    '"                        YDT_TBL t5a2" & vbCrLf & _
    '"                    ON  t5a1.YOYAKU_NO = t5a2.YOYAKU_NO" & vbCrLf & _
    '"                WHERE" & vbCrLf & _
    '"                    t5a1.YOYAKU_STS = '4'" & vbCrLf & _
    '"                AND t5a1.SHISETU_KBN = '1'" & vbCrLf & _
    '"                AND t5a1.SAIJI_BUNRUI = '1'" & vbCrLf & _
    '"                AND TO_DATE(t5a2.YOYAKU_DT, 'YYYY/MM/DD') BETWEEN CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' AS DATE)    " & vbCrLf & _
    '"                AND CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' as date) + cast('1 Month' as interval) - cast('1 day' as interval)" & vbCrLf & _
    '"                GROUP BY" & vbCrLf & _
    '"                    t5a2.YOYAKU_DT" & vbCrLf & _
    '"           ) TMP" & vbCrLf & _
    '"         ) t5a" & vbCrLf & _
    '"    CROSS JOIN" & vbCrLf & _
    '"        (" & vbCrLf & _
    '"          SELECT" & vbCrLf & _
    '"           COUNT(TMP.YOYAKU_DT) AS COUNT" & vbCrLf & _
    '"           FROM" & vbCrLf & _
    '"           (" & vbCrLf & _
    '"                SELECT" & vbCrLf & _
    '"                    t5b2.YOYAKU_DT" & vbCrLf & _
    '"                FROM" & vbCrLf & _
    '"                    YOYAKU_TBL t5b1" & vbCrLf & _
    '"                    LEFT JOIN" & vbCrLf & _
    '"                        YDT_TBL t5b2" & vbCrLf & _
    '"                    ON  t5b1.YOYAKU_NO = t5b2.YOYAKU_NO" & vbCrLf & _
    '"                WHERE" & vbCrLf & _
    '"                    t5b1.YOYAKU_STS = '4'" & vbCrLf & _
    '"                AND t5b1.SHISETU_KBN = '1'" & vbCrLf & _
    '"                AND t5b1.SAIJI_BUNRUI = '2'" & vbCrLf & _
    '"                AND TO_DATE(t5b2.YOYAKU_DT, 'YYYY/MM/DD') BETWEEN CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' AS DATE)    " & vbCrLf & _
    '"                AND CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' as date) + cast('1 Month' as interval) - cast('1 day' as interval)" & vbCrLf & _
    '"                GROUP BY" & vbCrLf & _
    '"                    t5b2.YOYAKU_DT" & vbCrLf & _
    '"           ) TMP" & vbCrLf & _
    '"         ) t5b" & vbCrLf & _
    '"    CROSS JOIN" & vbCrLf & _
    '"        (" & vbCrLf & _
    '"          SELECT" & vbCrLf & _
    '"           COUNT(TMP.YOYAKU_DT) AS COUNT" & vbCrLf & _
    '"           FROM" & vbCrLf & _
    '"           (" & vbCrLf & _
    '"                SELECT" & vbCrLf & _
    '"                    t5c2.YOYAKU_DT" & vbCrLf & _
    '"                FROM" & vbCrLf & _
    '"                    YOYAKU_TBL t5c1" & vbCrLf & _
    '"                    LEFT JOIN" & vbCrLf & _
    '"                        YDT_TBL t5c2" & vbCrLf & _
    '"                    ON  t5c1.YOYAKU_NO = t5c2.YOYAKU_NO" & vbCrLf & _
    '"                WHERE" & vbCrLf & _
    '"                    t5c1.YOYAKU_STS = '4'" & vbCrLf & _
    '"                AND t5c1.SHISETU_KBN = '1'" & vbCrLf & _
    '"                AND t5c1.SAIJI_BUNRUI = '3'" & vbCrLf & _
    '"                AND TO_DATE(t5c2.YOYAKU_DT, 'YYYY/MM/DD') BETWEEN CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' AS DATE)    " & vbCrLf & _
    '"                AND CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' as date) + cast('1 Month' as interval) - cast('1 day' as interval)" & vbCrLf & _
    '"                GROUP BY" & vbCrLf & _
    '"                    t5c2.YOYAKU_DT" & vbCrLf & _
    '"           ) TMP" & vbCrLf & _
    '"         ) t5c" & vbCrLf & _
    '"    CROSS JOIN" & vbCrLf & _
    '"        (" & vbCrLf & _
    '"          SELECT" & vbCrLf & _
    '"           COUNT(TMP.YOYAKU_DT) AS COUNT" & vbCrLf & _
    '"           FROM" & vbCrLf & _
    '"           (" & vbCrLf & _
    '"                SELECT" & vbCrLf & _
    '"                    t5d2.YOYAKU_DT" & vbCrLf & _
    '"                FROM" & vbCrLf & _
    '"                    YOYAKU_TBL t5d1" & vbCrLf & _
    '"                    LEFT JOIN" & vbCrLf & _
    '"                        YDT_TBL t5d2" & vbCrLf & _
    '"                    ON  t5d1.YOYAKU_NO = t5d2.YOYAKU_NO" & vbCrLf & _
    '"                WHERE" & vbCrLf & _
    '"                    t5d1.YOYAKU_STS = '4'" & vbCrLf & _
    '"                AND t5d1.SHISETU_KBN = '1'" & vbCrLf & _
    '"                AND t5d1.SAIJI_BUNRUI = '4'" & vbCrLf & _
    '"                AND TO_DATE(t5d2.YOYAKU_DT, 'YYYY/MM/DD') BETWEEN CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' AS DATE)    " & vbCrLf & _
    '"                AND CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' as date) + cast('1 Month' as interval) - cast('1 day' as interval)" & vbCrLf & _
    '"                GROUP BY" & vbCrLf & _
    '"                    t5d2.YOYAKU_DT" & vbCrLf & _
    '"           ) TMP" & vbCrLf & _
    '"         ) t5d" & vbCrLf & _
    '"    CROSS JOIN" & vbCrLf & _
    '"        (" & vbCrLf & _
    '"          SELECT" & vbCrLf & _
    '"           COUNT(TMP.YOYAKU_DT) AS COUNT" & vbCrLf & _
    '"           FROM" & vbCrLf & _
    '"           (" & vbCrLf & _
    '"                SELECT" & vbCrLf & _
    '"                    t5e2.YOYAKU_DT" & vbCrLf & _
    '"                FROM" & vbCrLf & _
    '"                    YOYAKU_TBL t5e1" & vbCrLf & _
    '"                    LEFT JOIN" & vbCrLf & _
    '"                        YDT_TBL t5e2" & vbCrLf & _
    '"                    ON  t5e1.YOYAKU_NO = t5e2.YOYAKU_NO" & vbCrLf & _
    '"                WHERE" & vbCrLf & _
    '"                    t5e1.YOYAKU_STS = '4'" & vbCrLf & _
    '"                AND t5e1.SHISETU_KBN = '1'" & vbCrLf & _
    '"                AND t5e1.SAIJI_BUNRUI = '5'" & vbCrLf & _
    '"                AND TO_DATE(t5e2.YOYAKU_DT, 'YYYY/MM/DD') BETWEEN CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' AS DATE)    " & vbCrLf & _
    '"                AND CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' as date) + cast('1 Month' as interval) - cast('1 day' as interval)" & vbCrLf & _
    '"                GROUP BY" & vbCrLf & _
    '"                    t5e2.YOYAKU_DT" & vbCrLf & _
    '"           ) TMP" & vbCrLf & _
    '"         ) t5e" & vbCrLf & _
    '"    CROSS JOIN" & vbCrLf & _
    '"        (" & vbCrLf & _
    '"          SELECT" & vbCrLf & _
    '"           COUNT(TMP.YOYAKU_DT) AS COUNT" & vbCrLf & _
    '"           FROM" & vbCrLf & _
    '"           (" & vbCrLf & _
    '"                SELECT" & vbCrLf & _
    '"                    t5f2.YOYAKU_DT" & vbCrLf & _
    '"                FROM" & vbCrLf & _
    '"                    YOYAKU_TBL t5f1" & vbCrLf & _
    '"                    LEFT JOIN" & vbCrLf & _
    '"                        YDT_TBL t5f2" & vbCrLf & _
    '"                    ON  t5f1.YOYAKU_NO = t5f2.YOYAKU_NO" & vbCrLf & _
    '"                WHERE" & vbCrLf & _
    '"                    t5f1.YOYAKU_STS = '4'" & vbCrLf & _
    '"                AND t5f1.SHISETU_KBN = '1'" & vbCrLf & _
    '"                AND t5f1.SAIJI_BUNRUI = '6'" & vbCrLf & _
    '"                AND TO_DATE(t5f2.YOYAKU_DT, 'YYYY/MM/DD') BETWEEN CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' AS DATE)    " & vbCrLf & _
    '"                AND CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' as date) + cast('1 Month' as interval) - cast('1 day' as interval)" & vbCrLf & _
    '"                GROUP BY" & vbCrLf & _
    '"                    t5f2.YOYAKU_DT" & vbCrLf & _
    '"           ) TMP" & vbCrLf & _
    '"        ) t5f"
    Private strSelectRiyoJokyoTheatreSQL As String = _
"    SELECT" & vbCrLf & _
"        EXTRACT(DAY FROM (CAST(:HYOJI_NEN ||:HYOJI_TSUKI || '01' as date) + cast('1 Month' as interval) - cast('1 day' as interval)))" & vbCrLf & _
"         - t1.COUNT - t2.COUNT AS RIYO_KANO_NISSU," & vbCrLf & _
"        t1.COUNT AS KYUKAN_NISSU," & vbCrLf & _
"        t2.COUNT AS MAINTE_NISSU," & vbCrLf & _
"        t4a.COUNT AS RIYO_NISSU_CHAKUSEKI," & vbCrLf & _
"        t4b.COUNT AS RIYO_NISSU_STANDING," & vbCrLf & _
"        t4c.COUNT AS RIYO_NISSU_MIX," & vbCrLf & _
"        t4d.COUNT AS RIYO_NISSU_SAIJI," & vbCrLf & _
"        t5a.COUNT AS RIYO_NISSU_MUSIC," & vbCrLf & _
"        t5b.COUNT AS RIYO_NISSU_ENGEKI," & vbCrLf & _
"        t5c.COUNT AS RIYO_NISSU_ENGEI," & vbCrLf & _
"        t5d.COUNT AS RIYO_NISSU_BUSINESS," & vbCrLf & _
"        t5e.COUNT AS RIYO_NISSU_MOVIE," & vbCrLf & _
"        t5f.COUNT AS RIYO_NISSU_SONOTA" & vbCrLf & _
"    FROM" & vbCrLf & _
"        (" & vbCrLf & _
"          SELECT" & vbCrLf & _
"           COUNT(TMP.HOLMENT_DT) AS COUNT" & vbCrLf & _
"           FROM" & vbCrLf & _
"           (" & vbCrLf & _
"            SELECT" & vbCrLf & _
"                HOLMENT_DT" & vbCrLf & _
"            FROM" & vbCrLf & _
"                HOLMENT_MST" & vbCrLf & _
"            WHERE" & vbCrLf & _
"            SHISETU_KBN = '1'" & vbCrLf & _
"            AND HOLMENT_KBN = '1'" & vbCrLf & _
"            AND TO_DATE(HOLMENT_DT, 'YYYY/MM/DD') BETWEEN CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' AS DATE)    " & vbCrLf & _
"                AND CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' as date) + cast('1 Month' as interval) - cast('1 day' as interval)" & vbCrLf & _
"            GROUP BY HOLMENT_DT" & vbCrLf & _
"           ) TMP" & vbCrLf & _
"        ) t1" & vbCrLf & _
"    CROSS JOIN" & vbCrLf & _
"        (" & vbCrLf & _
"          SELECT" & vbCrLf & _
"           COUNT(TMP.HOLMENT_DT) AS COUNT" & vbCrLf & _
"           FROM" & vbCrLf & _
"           (" & vbCrLf & _
"            SELECT" & vbCrLf & _
"                HOLMENT_DT" & vbCrLf & _
"            FROM" & vbCrLf & _
"                HOLMENT_MST" & vbCrLf & _
"            WHERE" & vbCrLf & _
"            SHISETU_KBN = '1'" & vbCrLf & _
"            AND HOLMENT_KBN = '2'" & vbCrLf & _
"            AND TO_DATE(HOLMENT_DT, 'YYYY/MM/DD') BETWEEN CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' AS DATE)    " & vbCrLf & _
"                AND CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' as date) + cast('1 Month' as interval) - cast('1 day' as interval)" & vbCrLf & _
"            GROUP BY HOLMENT_DT" & vbCrLf & _
"           ) TMP" & vbCrLf & _
"        ) t2" & vbCrLf & _
"    CROSS JOIN" & vbCrLf & _
"        (" & vbCrLf & _
"          SELECT" & vbCrLf & _
"           COUNT(TMP.YOYAKU_DT) AS COUNT" & vbCrLf & _
"           FROM" & vbCrLf & _
"           (" & vbCrLf & _
"            SELECT" & vbCrLf & _
"                t4a2.YOYAKU_DT" & vbCrLf & _
"            FROM" & vbCrLf & _
"                YOYAKU_TBL t4a1" & vbCrLf & _
"                LEFT JOIN" & vbCrLf & _
"                    YDT_TBL t4a2" & vbCrLf & _
"                ON  t4a1.YOYAKU_NO = t4a2.YOYAKU_NO" & vbCrLf & _
"            WHERE" & vbCrLf & _
"                t4a1.YOYAKU_STS = '4'" & vbCrLf & _
"            AND t4a1.SHISETU_KBN = '1'" & vbCrLf & _
"            AND t4a1.RIYO_TYPE = '1'" & vbCrLf & _
"            AND TO_DATE(t4a2.YOYAKU_DT, 'YYYY/MM/DD') BETWEEN CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' AS DATE)    " & vbCrLf & _
"                AND CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' as date) + cast('1 Month' as interval) - cast('1 day' as interval)" & vbCrLf & _
"            GROUP BY" & vbCrLf & _
"                t4a2.YOYAKU_DT" & vbCrLf & _
"           ) TMP" & vbCrLf & _
"        ) t4a" & vbCrLf & _
"    CROSS JOIN" & vbCrLf & _
"        (" & vbCrLf & _
"          SELECT" & vbCrLf & _
"           COUNT(TMP.YOYAKU_DT) AS COUNT" & vbCrLf & _
"           FROM" & vbCrLf & _
"           (" & vbCrLf & _
"                SELECT" & vbCrLf & _
"                    t4b2.YOYAKU_DT" & vbCrLf & _
"                FROM" & vbCrLf & _
"                    YOYAKU_TBL t4b1" & vbCrLf & _
"                    LEFT JOIN" & vbCrLf & _
"                        YDT_TBL t4b2" & vbCrLf & _
"                    ON  t4b1.YOYAKU_NO = t4b2.YOYAKU_NO" & vbCrLf & _
"                WHERE" & vbCrLf & _
"                    t4b1.YOYAKU_STS = '4'" & vbCrLf & _
"                AND t4b1.SHISETU_KBN = '1'" & vbCrLf & _
"                AND t4b1.RIYO_TYPE = '2'" & vbCrLf & _
"                AND TO_DATE(t4b2.YOYAKU_DT, 'YYYY/MM/DD') BETWEEN CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' AS DATE)    " & vbCrLf & _
"                AND CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' as date) + cast('1 Month' as interval) - cast('1 day' as interval)" & vbCrLf & _
"                GROUP BY" & vbCrLf & _
"                    t4b2.YOYAKU_DT" & vbCrLf & _
"           ) TMP" & vbCrLf & _
"        ) t4b" & vbCrLf & _
"    CROSS JOIN" & vbCrLf & _
"        (" & vbCrLf & _
"          SELECT" & vbCrLf & _
"           COUNT(TMP.YOYAKU_DT) AS COUNT" & vbCrLf & _
"           FROM" & vbCrLf & _
"           (" & vbCrLf & _
"                SELECT" & vbCrLf & _
"                    t4c2.YOYAKU_DT" & vbCrLf & _
"                FROM" & vbCrLf & _
"                    YOYAKU_TBL t4c1" & vbCrLf & _
"                    LEFT JOIN" & vbCrLf & _
"                        YDT_TBL t4c2" & vbCrLf & _
"                    ON  t4c1.YOYAKU_NO = t4c2.YOYAKU_NO" & vbCrLf & _
"                WHERE" & vbCrLf & _
"                    t4c1.YOYAKU_STS = '4'" & vbCrLf & _
"                AND t4c1.SHISETU_KBN = '1'" & vbCrLf & _
"                AND t4c1.RIYO_TYPE = '3'" & vbCrLf & _
"                AND TO_DATE(t4c2.YOYAKU_DT, 'YYYY/MM/DD') BETWEEN CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' AS DATE)    " & vbCrLf & _
"                AND CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' as date) + cast('1 Month' as interval) - cast('1 day' as interval)" & vbCrLf & _
"                GROUP BY" & vbCrLf & _
"                    t4c2.YOYAKU_DT" & vbCrLf & _
"           ) TMP" & vbCrLf & _
"        ) t4c" & vbCrLf & _
"    CROSS JOIN" & vbCrLf & _
"        (" & vbCrLf & _
"          SELECT" & vbCrLf & _
"           COUNT(TMP.YOYAKU_DT) AS COUNT" & vbCrLf & _
"           FROM" & vbCrLf & _
"           (" & vbCrLf & _
"                SELECT" & vbCrLf & _
"                    t4d2.YOYAKU_DT" & vbCrLf & _
"                FROM" & vbCrLf & _
"                    YOYAKU_TBL t4d1" & vbCrLf & _
"                    LEFT JOIN" & vbCrLf & _
"                        YDT_TBL t4d2" & vbCrLf & _
"                    ON  t4d1.YOYAKU_NO = t4d2.YOYAKU_NO" & vbCrLf & _
"                WHERE" & vbCrLf & _
"                    t4d1.YOYAKU_STS = '4'" & vbCrLf & _
"                AND t4d1.SHISETU_KBN = '1'" & vbCrLf & _
"                AND t4d1.RIYO_TYPE = '4'" & vbCrLf & _
"                AND TO_DATE(t4d2.YOYAKU_DT, 'YYYY/MM/DD') BETWEEN CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' AS DATE)    " & vbCrLf & _
"                AND CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' as date) + cast('1 Month' as interval) - cast('1 day' as interval)" & vbCrLf & _
"                GROUP BY" & vbCrLf & _
"                    t4d2.YOYAKU_DT" & vbCrLf & _
"           ) TMP" & vbCrLf & _
"        ) t4d" & vbCrLf & _
"    CROSS JOIN" & vbCrLf & _
"        (" & vbCrLf & _
"          SELECT" & vbCrLf & _
"           COUNT(TMP.YOYAKU_DT) AS COUNT" & vbCrLf & _
"           FROM" & vbCrLf & _
"           (" & vbCrLf & _
"                SELECT" & vbCrLf & _
"                    t5a2.YOYAKU_DT" & vbCrLf & _
"                FROM" & vbCrLf & _
"                    YOYAKU_TBL t5a1" & vbCrLf & _
"                    LEFT JOIN" & vbCrLf & _
"                        YDT_TBL t5a2" & vbCrLf & _
"                    ON  t5a1.YOYAKU_NO = t5a2.YOYAKU_NO" & vbCrLf & _
"                WHERE" & vbCrLf & _
"                    t5a1.YOYAKU_STS = '4'" & vbCrLf & _
"                AND t5a1.SHISETU_KBN = '1'" & vbCrLf & _
"                AND t5a1.SAIJI_BUNRUI = '1'" & vbCrLf & _
"                AND TO_DATE(t5a2.YOYAKU_DT, 'YYYY/MM/DD') BETWEEN CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' AS DATE)    " & vbCrLf & _
"                AND CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' as date) + cast('1 Month' as interval) - cast('1 day' as interval)" & vbCrLf & _
"                GROUP BY" & vbCrLf & _
"                    t5a2.YOYAKU_DT" & vbCrLf & _
"           ) TMP" & vbCrLf & _
"         ) t5a" & vbCrLf & _
"    CROSS JOIN" & vbCrLf & _
"        (" & vbCrLf & _
"          SELECT" & vbCrLf & _
"           COUNT(TMP.YOYAKU_DT) AS COUNT" & vbCrLf & _
"           FROM" & vbCrLf & _
"           (" & vbCrLf & _
"                SELECT" & vbCrLf & _
"                    t5b2.YOYAKU_DT" & vbCrLf & _
"                FROM" & vbCrLf & _
"                    YOYAKU_TBL t5b1" & vbCrLf & _
"                    LEFT JOIN" & vbCrLf & _
"                        YDT_TBL t5b2" & vbCrLf & _
"                    ON  t5b1.YOYAKU_NO = t5b2.YOYAKU_NO" & vbCrLf & _
"                WHERE" & vbCrLf & _
"                    t5b1.YOYAKU_STS = '4'" & vbCrLf & _
"                AND t5b1.SHISETU_KBN = '1'" & vbCrLf & _
"                AND t5b1.SAIJI_BUNRUI = '2'" & vbCrLf & _
"                AND TO_DATE(t5b2.YOYAKU_DT, 'YYYY/MM/DD') BETWEEN CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' AS DATE)    " & vbCrLf & _
"                AND CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' as date) + cast('1 Month' as interval) - cast('1 day' as interval)" & vbCrLf & _
"                GROUP BY" & vbCrLf & _
"                    t5b2.YOYAKU_DT" & vbCrLf & _
"           ) TMP" & vbCrLf & _
"         ) t5b" & vbCrLf & _
"    CROSS JOIN" & vbCrLf & _
"        (" & vbCrLf & _
"          SELECT" & vbCrLf & _
"           COUNT(TMP.YOYAKU_DT) AS COUNT" & vbCrLf & _
"           FROM" & vbCrLf & _
"           (" & vbCrLf & _
"                SELECT" & vbCrLf & _
"                    t5c2.YOYAKU_DT" & vbCrLf & _
"                FROM" & vbCrLf & _
"                    YOYAKU_TBL t5c1" & vbCrLf & _
"                    LEFT JOIN" & vbCrLf & _
"                        YDT_TBL t5c2" & vbCrLf & _
"                    ON  t5c1.YOYAKU_NO = t5c2.YOYAKU_NO" & vbCrLf & _
"                WHERE" & vbCrLf & _
"                    t5c1.YOYAKU_STS = '4'" & vbCrLf & _
"                AND t5c1.SHISETU_KBN = '1'" & vbCrLf & _
"                AND t5c1.SAIJI_BUNRUI = '3'" & vbCrLf & _
"                AND TO_DATE(t5c2.YOYAKU_DT, 'YYYY/MM/DD') BETWEEN CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' AS DATE)    " & vbCrLf & _
"                AND CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' as date) + cast('1 Month' as interval) - cast('1 day' as interval)" & vbCrLf & _
"                GROUP BY" & vbCrLf & _
"                    t5c2.YOYAKU_DT" & vbCrLf & _
"           ) TMP" & vbCrLf & _
"         ) t5c" & vbCrLf & _
"    CROSS JOIN" & vbCrLf & _
"        (" & vbCrLf & _
"          SELECT" & vbCrLf & _
"           COUNT(TMP.YOYAKU_DT) AS COUNT" & vbCrLf & _
"           FROM" & vbCrLf & _
"           (" & vbCrLf & _
"                SELECT" & vbCrLf & _
"                    t5d2.YOYAKU_DT" & vbCrLf & _
"                FROM" & vbCrLf & _
"                    YOYAKU_TBL t5d1" & vbCrLf & _
"                    LEFT JOIN" & vbCrLf & _
"                        YDT_TBL t5d2" & vbCrLf & _
"                    ON  t5d1.YOYAKU_NO = t5d2.YOYAKU_NO" & vbCrLf & _
"                WHERE" & vbCrLf & _
"                    t5d1.YOYAKU_STS = '4'" & vbCrLf & _
"                AND t5d1.SHISETU_KBN = '1'" & vbCrLf & _
"                AND t5d1.SAIJI_BUNRUI = '4'" & vbCrLf & _
"                AND TO_DATE(t5d2.YOYAKU_DT, 'YYYY/MM/DD') BETWEEN CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' AS DATE)    " & vbCrLf & _
"                AND CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' as date) + cast('1 Month' as interval) - cast('1 day' as interval)" & vbCrLf & _
"                GROUP BY" & vbCrLf & _
"                    t5d2.YOYAKU_DT" & vbCrLf & _
"           ) TMP" & vbCrLf & _
"         ) t5d" & vbCrLf & _
"    CROSS JOIN" & vbCrLf & _
"        (" & vbCrLf & _
"          SELECT" & vbCrLf & _
"           COUNT(TMP.YOYAKU_DT) AS COUNT" & vbCrLf & _
"           FROM" & vbCrLf & _
"           (" & vbCrLf & _
"                SELECT" & vbCrLf & _
"                    t5e2.YOYAKU_DT" & vbCrLf & _
"                FROM" & vbCrLf & _
"                    YOYAKU_TBL t5e1" & vbCrLf & _
"                    LEFT JOIN" & vbCrLf & _
"                        YDT_TBL t5e2" & vbCrLf & _
"                    ON  t5e1.YOYAKU_NO = t5e2.YOYAKU_NO" & vbCrLf & _
"                WHERE" & vbCrLf & _
"                    t5e1.YOYAKU_STS = '4'" & vbCrLf & _
"                AND t5e1.SHISETU_KBN = '1'" & vbCrLf & _
"                AND t5e1.SAIJI_BUNRUI = '5'" & vbCrLf & _
"                AND TO_DATE(t5e2.YOYAKU_DT, 'YYYY/MM/DD') BETWEEN CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' AS DATE)    " & vbCrLf & _
"                AND CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' as date) + cast('1 Month' as interval) - cast('1 day' as interval)" & vbCrLf & _
"                GROUP BY" & vbCrLf & _
"                    t5e2.YOYAKU_DT" & vbCrLf & _
"           ) TMP" & vbCrLf & _
"         ) t5e" & vbCrLf & _
"    CROSS JOIN" & vbCrLf & _
"        (" & vbCrLf & _
"          SELECT" & vbCrLf & _
"           COUNT(TMP.YOYAKU_DT) AS COUNT" & vbCrLf & _
"           FROM" & vbCrLf & _
"           (" & vbCrLf & _
"                SELECT" & vbCrLf & _
"                    t5f2.YOYAKU_DT" & vbCrLf & _
"                FROM" & vbCrLf & _
"                    YOYAKU_TBL t5f1" & vbCrLf & _
"                    LEFT JOIN" & vbCrLf & _
"                        YDT_TBL t5f2" & vbCrLf & _
"                    ON  t5f1.YOYAKU_NO = t5f2.YOYAKU_NO" & vbCrLf & _
"                WHERE" & vbCrLf & _
"                    t5f1.YOYAKU_STS = '4'" & vbCrLf & _
"                AND t5f1.SHISETU_KBN = '1'" & vbCrLf & _
"                AND t5f1.SAIJI_BUNRUI = '6'" & vbCrLf & _
"                AND TO_DATE(t5f2.YOYAKU_DT, 'YYYY/MM/DD') BETWEEN CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' AS DATE)    " & vbCrLf & _
"                AND CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' as date) + cast('1 Month' as interval) - cast('1 day' as interval)" & vbCrLf & _
"                GROUP BY" & vbCrLf & _
"                    t5f2.YOYAKU_DT" & vbCrLf & _
"           ) TMP" & vbCrLf & _
"        ) t5f"
    ' 2017.07.31 e.watanabe UPD END↑ シアター利用形状の修正（1:利用日数<着席>、2:利用日数<スタンディング>）
#End Region

#Region "SQL文(利用状況一覧(スタジオ)取得)"

    'SQL文(利用状況一覧(スタジオ)取得)
    Private strSelectRiyoJokyoStudioSQL As String = _
    "    SELECT" & vbCrLf & _
"        EXTRACT(DAY FROM (CAST(:HYOJI_NEN ||:HYOJI_TSUKI || '01' as date) + cast('1 Month' as interval) - cast('1 day' as interval)))" & vbCrLf & _
"         - t1.COUNT - t2.COUNT AS RIYO_KANO_NISSU," & vbCrLf & _
"        t1.COUNT AS KYUKAN_NISSU," & vbCrLf & _
"        t2.COUNT AS MAINTE_NISSU," & vbCrLf & _
"        t4a.COUNT AS RIYO_NISSU_201st," & vbCrLf & _
"        t4b.COUNT AS RIYO_NISSU_202st," & vbCrLf & _
"        t4c.COUNT AS RIYO_NISSU_houselock" & vbCrLf & _
"    FROM" & vbCrLf & _
"        (" & vbCrLf & _
"        SELECT" & vbCrLf & _
"          COUNT(TMP.HOLMENT_DT) AS COUNT" & vbCrLf & _
"           FROM" & vbCrLf & _
"            (" & vbCrLf & _
"            SELECT" & vbCrLf & _
"                HOLMENT_DT" & vbCrLf & _
"            FROM" & vbCrLf & _
"                HOLMENT_MST" & vbCrLf & _
"            WHERE" & vbCrLf & _
"            SHISETU_KBN = '2'" & vbCrLf & _
"            AND HOLMENT_KBN = '1'" & vbCrLf & _
"            AND TO_DATE(HOLMENT_DT, 'YYYY/MM/DD') BETWEEN CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' AS DATE)" & vbCrLf & _
"                AND CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' as date) + cast('1 Month' as interval) - cast('1 day' as interval)" & vbCrLf & _
"            GROUP BY HOLMENT_DT" & vbCrLf & _
"           ) TMP" & vbCrLf & _
"        ) t1" & vbCrLf & _
"    CROSS JOIN" & vbCrLf & _
"        (" & vbCrLf & _
"        SELECT" & vbCrLf & _
"          COUNT(TMP.HOLMENT_DT) AS COUNT" & vbCrLf & _
"           FROM" & vbCrLf & _
"            (" & vbCrLf & _
"            SELECT" & vbCrLf & _
"                HOLMENT_DT" & vbCrLf & _
"            FROM" & vbCrLf & _
"                HOLMENT_MST" & vbCrLf & _
"            WHERE" & vbCrLf & _
"            SHISETU_KBN = '2'            " & vbCrLf & _
"            AND HOLMENT_KBN = '2'" & vbCrLf & _
"            AND TO_DATE(HOLMENT_DT, 'YYYY/MM/DD') BETWEEN CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' AS DATE)" & vbCrLf & _
"                AND CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' as date) + cast('1 Month' as interval) - cast('1 day' as interval)" & vbCrLf & _
"            GROUP BY HOLMENT_DT" & vbCrLf & _
"           ) TMP" & vbCrLf & _
"        ) t2" & vbCrLf & _
"    CROSS JOIN" & vbCrLf & _
"            (" & vbCrLf & _
"            SELECT" & vbCrLf & _
"              COUNT(TMP.YOYAKU_DT) AS COUNT" & vbCrLf & _
"               FROM" & vbCrLf & _
"                (" & vbCrLf & _
"                SELECT" & vbCrLf & _
"                    t4a2.YOYAKU_DT" & vbCrLf & _
"                FROM" & vbCrLf & _
"                    YOYAKU_TBL t4a1" & vbCrLf & _
"                    LEFT JOIN" & vbCrLf & _
"                        YDT_TBL t4a2" & vbCrLf & _
"                    ON  t4a1.YOYAKU_NO = t4a2.YOYAKU_NO" & vbCrLf & _
"                WHERE" & vbCrLf & _
"                    t4a1.YOYAKU_STS = '4'" & vbCrLf & _
"                AND t4a1.SHISETU_KBN = '2'" & vbCrLf & _
"                AND t4a1.STUDIO_KBN = '1'" & vbCrLf & _
"                AND TO_DATE(t4a2.YOYAKU_DT, 'YYYY/MM/DD') BETWEEN CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' AS DATE)" & vbCrLf & _
"                AND CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' as date) + cast('1 Month' as interval) - cast('1 day' as interval)" & vbCrLf & _
"                GROUP BY" & vbCrLf & _
"                    t4a2.YOYAKU_DT" & vbCrLf & _
"           ) TMP" & vbCrLf & _
"         ) t4a" & vbCrLf & _
"    CROSS JOIN" & vbCrLf & _
"            (" & vbCrLf & _
"            SELECT" & vbCrLf & _
"              COUNT(TMP.YOYAKU_DT) AS COUNT" & vbCrLf & _
"               FROM" & vbCrLf & _
"                (" & vbCrLf & _
"                SELECT" & vbCrLf & _
"                    t4b2.YOYAKU_DT" & vbCrLf & _
"                FROM" & vbCrLf & _
"                    YOYAKU_TBL t4b1" & vbCrLf & _
"                    LEFT JOIN" & vbCrLf & _
"                        YDT_TBL t4b2" & vbCrLf & _
"                    ON  t4b1.YOYAKU_NO = t4b2.YOYAKU_NO" & vbCrLf & _
"                WHERE" & vbCrLf & _
"                    t4b1.YOYAKU_STS = '4'" & vbCrLf & _
"                AND t4b1.SHISETU_KBN = '2'" & vbCrLf & _
"                AND t4b1.STUDIO_KBN = '2'" & vbCrLf & _
"                AND TO_DATE(t4b2.YOYAKU_DT, 'YYYY/MM/DD') BETWEEN CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' AS DATE)" & vbCrLf & _
"                AND CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' as date) + cast('1 Month' as interval) - cast('1 day' as interval)" & vbCrLf & _
"                GROUP BY" & vbCrLf & _
"                    t4b2.YOYAKU_DT" & vbCrLf & _
"           ) TMP" & vbCrLf & _
"         ) t4b" & vbCrLf & _
"    CROSS JOIN" & vbCrLf & _
"            (" & vbCrLf & _
"            SELECT" & vbCrLf & _
"              COUNT(TMP.YOYAKU_DT) AS COUNT" & vbCrLf & _
"               FROM" & vbCrLf & _
"                (" & vbCrLf & _
"                SELECT" & vbCrLf & _
"                    t4c2.YOYAKU_DT" & vbCrLf & _
"                FROM" & vbCrLf & _
"                    YOYAKU_TBL t4c1" & vbCrLf & _
"                    LEFT JOIN" & vbCrLf & _
"                        YDT_TBL t4c2" & vbCrLf & _
"                    ON  t4c1.YOYAKU_NO = t4c2.YOYAKU_NO" & vbCrLf & _
"                WHERE" & vbCrLf & _
"                    t4c1.YOYAKU_STS = '4'" & vbCrLf & _
"                AND t4c1.SHISETU_KBN = '2'" & vbCrLf & _
"                AND t4c1.STUDIO_KBN = '3'" & vbCrLf & _
"                AND TO_DATE(t4c2.YOYAKU_DT, 'YYYY/MM/DD') BETWEEN CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' AS DATE)" & vbCrLf & _
"                AND CAST(:HYOJI_NEN || :HYOJI_TSUKI || '01' as date) + cast('1 Month' as interval) - cast('1 day' as interval)" & vbCrLf & _
"                GROUP BY" & vbCrLf & _
"                    t4c2.YOYAKU_DT" & vbCrLf & _
"           ) TMP" & vbCrLf & _
"        ) t4c"

#End Region

    '*****条件追加・変更、アダプタ・パラメータ網羅によるSQL文*****
#Region "利用状況一覧取得"

    ''' <summary>
    ''' 利用状況一覧取得
    ''' </summary>
    ''' <param name="Adapter">アダプタ</param>
    ''' <param name="Cn">コネクション</param>
    ''' <param name="ShisetuKbn">施設区分</param>
    ''' <param name="HyojiNen">表示年</param>
    ''' <param name="HyojiTsuki">表示月</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>利用状況一覧を取得するSQLの作成
    ''' <para>作成情報：2015/09/24 h.endo
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function setSelectRiyoJokyo(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByRef Cn As NpgsqlConnection, _
                                       ByVal ShisetuKbn As String, _
                                       ByVal HyojiNen As String, _
                                       ByVal HyojiTsuki As String) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = String.Empty

        Try

            'SQL文(SELECT)
            'データアダプタに、SQLのSELECT文を設定
            If ShisetuKbn = SHISETU_KBN_THEATER Then
                'シアターの場合
                strSQL = strSelectRiyoJokyoTheatreSQL
            Else
                'スタジオの場合
                strSQL = strSelectRiyoJokyoStudioSQL
            End If

            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            '条件項目の型を指定
            'バインド変数の型を設定
            With Adapter.SelectCommand.Parameters
                .Add(New NpgsqlParameter(":HYOJI_NEN", NpgsqlTypes.NpgsqlDbType.Varchar))      '表示年
                .Add(New NpgsqlParameter(":HYOJI_TSUKI", NpgsqlTypes.NpgsqlDbType.Varchar))    '表示月
            End With

            'バインド変数の値を設定
            With Adapter.SelectCommand
                .Parameters(":HYOJI_NEN").Value = HyojiNen          '表示年
                .Parameters(":HYOJI_TSUKI").Value = HyojiTsuki      '表示月
            End With

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
