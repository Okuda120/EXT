﻿Imports Common
Imports CommonEXT
Imports Npgsql
Imports System.Text

Public Class SqlEXTC0103

    Private commonLogicEXT As New CommonLogicEXT        '共通クラス

    'SQL文(請求情報)
    ' 2016.02.03 UPD START↓ h.hagiwara 登録済請求情報の更新判定取得追加
    'Private strEX05S002 As String = _
    '                        "SELECT " & vbCrLf & _
    '                        "    t1.yoyaku_no, " & vbCrLf & _
    '                        "    t1.seq, " & vbCrLf & _
    '                        "    t1.seikyu_irai_no, " & vbCrLf & _
    '                        "    t1.seikyu_dt, " & vbCrLf & _
    '                        "    t1.nyukin_yotei_dt, " & vbCrLf & _
    '                        "    t1.kakutei_kin, " & vbCrLf & _
    '                        "    t1.chosei_kin, " & vbCrLf & _
    '                        "    t1.shokei, " & vbCrLf & _
    '                        "    t1.tax_kin, " & vbCrLf & _
    '                        "    t1.seikyu_kin, " & vbCrLf & _
    '                        "    t1.seikyu_naiyo, " & vbCrLf & _
    '                        "    CASE WHEN t1.seikyu_naiyo='1' THEN '利用料' WHEN t1.seikyu_naiyo='2' THEN '付帯設備'" & vbCrLf & _
    '                        "         WHEN t1.seikyu_naiyo='3' THEN '利用料+付帯設備' WHEN t1.seikyu_naiyo='4' THEN '還付' ELSE '' END as seikyu_naiyo_nm, " & vbCrLf & _
    '                        "    t1.aite_cd, " & vbCrLf & _
    '                        "    m1.aite_nm, " & vbCrLf & _
    '                        "    t1.nyukin_kbn, " & vbCrLf & _
    '                        "    t1.seikyu_title1, " & vbCrLf & _
    '                        "    t1.seikyu_title2, " & vbCrLf & _
    '                        "    t1.nyukin_dt, " & vbCrLf & _
    '                        "    t1.nyukin_kin, " & vbCrLf & _
    '                        "    t1.seikyu_input_flg, " & vbCrLf & _
    '                        "    t1.seikyu_irai_flg, " & vbCrLf & _
    '                        "    t1.nyukin_input_flg, " & vbCrLf & _
    '                        "    t1.nyukin_link_no" & vbCrLf & _
    '                        "FROM " & vbCrLf & _
    '                        "    BILLPAY_TBL t1 " & vbCrLf & _
    '                        "LEFT JOIN AITESAKI_MST m1 " & vbCrLf & _
    '                        "    ON t1.aite_cd = m1.aite_cd " & vbCrLf & _
    '                        "WHERE t1.yoyaku_no = :YoyakuNo " & vbCrLf & _
    '                        "ORDER BY t1.seq "
    Private strEX05S002 As String = _
                            "SELECT " & vbCrLf & _
                            "    t1.yoyaku_no, " & vbCrLf & _
                            "    t1.seq, " & vbCrLf & _
                            "    t1.seikyu_irai_no, " & vbCrLf & _
                            "    t1.seikyu_dt, " & vbCrLf & _
                            "    t1.nyukin_yotei_dt, " & vbCrLf & _
                            "    t1.kakutei_kin, " & vbCrLf & _
                            "    t1.chosei_kin, " & vbCrLf & _
                            "    t1.shokei, " & vbCrLf & _
                            "    t1.tax_kin, " & vbCrLf & _
                            "    t1.seikyu_kin, " & vbCrLf & _
                            "    t1.seikyu_naiyo, " & vbCrLf & _
                            "    CASE WHEN t1.seikyu_naiyo='1' THEN '利用料' WHEN t1.seikyu_naiyo='2' THEN '付帯設備'" & vbCrLf & _
                            "         WHEN t1.seikyu_naiyo='3' THEN '利用料+付帯設備' WHEN t1.seikyu_naiyo='4' THEN '還付' ELSE '' END as seikyu_naiyo_nm, " & vbCrLf & _
                            "    t1.aite_cd, " & vbCrLf & _
                            "    m1.aite_nm, " & vbCrLf & _
                            "    t1.nyukin_kbn, " & vbCrLf & _
                            "    t1.seikyu_title1, " & vbCrLf & _
                            "    t1.seikyu_title2, " & vbCrLf & _
                            "    t1.nyukin_dt, " & vbCrLf & _
                            "    t1.nyukin_kin, " & vbCrLf & _
                            "    t1.seikyu_input_flg, " & vbCrLf & _
                            "    t1.seikyu_irai_flg, " & vbCrLf & _
                            "    t1.nyukin_input_flg, " & vbCrLf & _
                            "    t1.nyukin_link_no," & vbCrLf & _
                            "    t1.seikyu_input_flg AS get_seikyu_input_flg " & vbCrLf & _
                            "FROM " & vbCrLf & _
                            "    BILLPAY_TBL t1 " & vbCrLf & _
                            "LEFT JOIN AITESAKI_MST m1 " & vbCrLf & _
                            "    ON t1.aite_cd = m1.aite_cd " & vbCrLf & _
                            "WHERE t1.yoyaku_no = :YoyakuNo " & vbCrLf & _
                            "ORDER BY t1.seq "
    ' 2016.02.03 UPD END↑ h.hagiwara 登録済請求情報の更新判定取得追加

    '付帯設備
    'Private strEX05S004 As String = _
    '                        "SELECT " & vbCrLf & _
    '                        "    yoyaku_no, " & vbCrLf & _
    '                        "    yoyaku_dt, " & vbCrLf & _
    '                        "    fuzoku_nm, " & vbCrLf & _
    '                        "    total_fuzoku_kin " & vbCrLf & _
    '                        "FROM " & vbCrLf & _
    '                        "    friyo_tbl " & vbCrLf & _
    '                        "WHERE yoyaku_no = :YoyakuNo " & vbCrLf & _
    '                        "    AND yoyaku_dt = :Riyobi " & vbCrLf & _
    '                        "ORDER BY add_dt "
    Private strEX05S004 As String = _
                            "SELECT " & vbCrLf & _
                            "    t1.yoyaku_no, " & vbCrLf & _
                            "    t1.yoyaku_dt, " & vbCrLf & _
                            "    t1.fuzoku_nm, " & vbCrLf & _
                            "    t1.total_fuzoku_kin " & vbCrLf & _
                            "   ,CASE WHEN :SYANAI = '2' THEN 0 ELSE COALESCE(t2.tax_kin,0) END AS tax_kin" & vbCrLf & _
                            "FROM " & vbCrLf & _
                            "    friyo_tbl t1 " & vbCrLf & _
                            "LEFT JOIN ( SELECT t0.YOYAKU_NO , t0.YOYAKU_DT , sum(t0.tax_kin) AS tax_kin " & vbCrLf & _
                                       " FROM  ( SELECT CASE WHEN tt2.NOTAX_FLG = '0' THEN ROUND((tt1.FUTAI_KIN * tt3.TAX_RITU ) /100 ,0 )" & vbCrLf & _
                                       "                                              ELSE 0 " & vbCrLf & _
                                       "                END TAX_KIN " & vbCrLf & _
                                       "               ,tt1.YOYAKU_NO " & vbCrLf & _
                                       "               ,tt1.YOYAKU_DT " & vbCrLf & _
                                       "         FROM  FRIYO_MEISAI_TBL tt1 " & vbCrLf & _
                                       "           LEFT OUTER JOIN  FBUNRUI_MST tt2 " & vbCrLf & _
                                       "                       ON   tt1.FUTAI_BUNRUI_CD = tt2.BUNRUI_CD " & vbCrLf & _
                                       "                      AND   '2' = tt2.SHISETU_KBN " & vbCrLf & _
                                       "                      AND   tt1.YOYAKU_DT BETWEEN tt2.KIKAN_FROM AND tt2.KIKAN_TO " & vbCrLf & _
                                       "           LEFT OUTER JOIN  TAX_MST tt3 " & vbCrLf & _
                                       "                       ON   tt1.YOYAKU_DT BETWEEN tt3.TAXS_DT AND tt3.TAXE_DT " & vbCrLf & _
                                       "        ) t0 " & vbCrLf & _
                                       " GROUP BY t0.YOYAKU_NO , t0.YOYAKU_DT " & vbCrLf & _
                            "       ) t2 " & vbCrLf & _
                            "      ON  t1.YOYAKU_NO = t2.YOYAKU_NO " & vbCrLf & _
                            "     AND  t1.YOYAKU_DT = t2.YOYAKU_DT " & vbCrLf & _
                            "WHERE t1.yoyaku_no = :YoyakuNo " & vbCrLf & _
                            " AND  t1.yoyaku_dt = :Riyobi " & vbCrLf & _
                            "ORDER BY add_dt "

    '付帯設備明細
    ' --- 2019/07/24 軽減税率対応 Start E.Okuda@Compass ---
    Private strEX05S098 As String =
                            "SELECT " & vbCrLf &
                            "    t1.yoyaku_no, " & vbCrLf &
                            "    t1.yoyaku_dt, " & vbCrLf &
                            "    t1.futai_bunrui_cd, " & vbCrLf &
                            "    t1.futai_cd, " & vbCrLf &
                            "    t1.futai_tanka, " & vbCrLf &
                            "    t1.futai_su, " & vbCrLf &
                            "    t1.futai_shokei, " & vbCrLf &
                            "    t1.futai_chosei, " & vbCrLf &
                            "    t1.futai_kin, " & vbCrLf &
                            "    t1.futai_biko, " & vbCrLf &
                            "    m1.futai_nm, " & vbCrLf &
                            "    m1.tani as futai_tani, " & vbCrLf &
                            "    m2.bunrui_nm as futai_bunrui_nm, " & vbCrLf &
                            "    m2.shukei_grp, " & vbCrLf &
                            "    m2.kamoku_cd, " & vbCrLf &
                            "    m2.saimoku_cd, " & vbCrLf &
                            "    m2.uchi_cd, " & vbCrLf &
                            "    m2.shosai_cd, " & vbCrLf &
                            "    m2.karikamoku_cd, " & vbCrLf &
                            "    m2.kari_saimoku_cd, " & vbCrLf &
                            "    m2.kari_uchi_cd, " & vbCrLf &
                            "    m2.kari_shosai_cd, " & vbCrLf &
                            "    m2.notax_flg, " & vbCrLf &
                            "    (CASE WHEN m2.zeiritsu IS NULL THEN m4.tax_ritu ELSE m2.zeiritsu END) as zeiritsu, " & vbCrLf &
                            "    m3.kamoku_nm, " & vbCrLf &
                            "    m3.saimoku_nm, " & vbCrLf &
                            "    m3.uchi_nm, " & vbCrLf &
                            "    m3.shosai_nm " & vbCrLf &
                            "   ,COALESCE((CASE WHEN :SYANAI = '2' THEN  0" & vbCrLf &
                            "                            ELSE " & vbCrLf &
                            "                               CASE WHEN m2.notax_flg = '0' THEN ROUND((t1.futai_kin * (CASE WHEN m2.zeiritsu IS NULL THEN m4.tax_ritu ELSE m2.zeiritsu END)) /100 ,0)" & vbCrLf &
                            "                                    ELSE 0 " & vbCrLf &
                            "                               END " & vbCrLf &
                            "     END),0)  as tax_kin " & vbCrLf &
                            "FROM " & vbCrLf &
                            "    friyo_meisai_tbl t1 " & vbCrLf &
                            "LEFT JOIN futai_mst m1 " & vbCrLf &
                            "    ON t1.futai_bunrui_cd = m1.bunrui_cd " & vbCrLf &
                            "    AND t1.futai_cd = m1.futai_cd " & vbCrLf &
                            "    AND m1.shisetu_kbn = :ShisetuKbn " & vbCrLf &
                            "    AND :Riyobi BETWEEN m1.kikan_from AND m1.kikan_to " & vbCrLf &
                            "    AND m1.sts = '0' " & vbCrLf &
                            "LEFT JOIN fbunrui_mst m2 " & vbCrLf &
                            "    ON t1.futai_bunrui_cd = m2.bunrui_cd " & vbCrLf &
                            "    AND m2.shisetu_kbn = :ShisetuKbn " & vbCrLf &
                            "    AND :Riyobi BETWEEN m2.kikan_from AND m2.kikan_to " & vbCrLf &
                            "    AND m2.sts = '0' " & vbCrLf &
                            "LEFT JOIN kamoku_mst m3 " & vbCrLf &
                            "    ON m2.kamoku_cd = m3.kamoku_cd " & vbCrLf &
                            "    AND m2.saimoku_cd = m3.saimoku_cd " & vbCrLf &
                            "    AND m2.uchi_cd = m3.uchi_cd " & vbCrLf &
                            "    AND m2.shosai_cd = m3.shosai_cd " & vbCrLf &
                            "    AND m3.sts = '0' " & vbCrLf &
                            "LEFT JOIN TAX_MST m4 " & vbCrLf &
                            "    ON  :Riyobi BETWEEN m4.TAXS_DT AND m4.TAXE_DT" & vbCrLf &
                            "WHERE " & vbCrLf &
                            "    t1.yoyaku_no = :YoyakuNo " & vbCrLf &
                            "    AND t1.yoyaku_dt = :Riyobi " & vbCrLf &
                            "ORDER BY t1.add_dt "

    'Private strEX05S098 As String =
    '                        "SELECT " & vbCrLf &
    '                        "    t1.yoyaku_no, " & vbCrLf &
    '                        "    t1.yoyaku_dt, " & vbCrLf &
    '                        "    t1.futai_bunrui_cd, " & vbCrLf &
    '                        "    t1.futai_cd, " & vbCrLf &
    '                        "    t1.futai_tanka, " & vbCrLf &
    '                        "    t1.futai_su, " & vbCrLf &
    '                        "    t1.futai_shokei, " & vbCrLf &
    '                        "    t1.futai_chosei, " & vbCrLf &
    '                        "    t1.futai_kin, " & vbCrLf &
    '                        "    t1.futai_biko, " & vbCrLf &
    '                        "    m1.futai_nm, " & vbCrLf &
    '                        "    m1.tani as futai_tani, " & vbCrLf &
    '                        "    m2.bunrui_nm as futai_bunrui_nm, " & vbCrLf &
    '                        "    m2.shukei_grp, " & vbCrLf &
    '                        "    m2.kamoku_cd, " & vbCrLf &
    '                        "    m2.saimoku_cd, " & vbCrLf &
    '                        "    m2.uchi_cd, " & vbCrLf &
    '                        "    m2.shosai_cd, " & vbCrLf &
    '                        "    m2.karikamoku_cd, " & vbCrLf &
    '                        "    m2.kari_saimoku_cd, " & vbCrLf &
    '                        "    m2.kari_uchi_cd, " & vbCrLf &
    '                        "    m2.kari_shosai_cd, " & vbCrLf &
    '                        "    m2.notax_flg, " & vbCrLf &
    '                        "    m3.kamoku_nm, " & vbCrLf &
    '                        "    m3.saimoku_nm, " & vbCrLf &
    '                        "    m3.uchi_nm, " & vbCrLf &
    '                        "    m3.shosai_nm " & vbCrLf &
    '                        "   ,COALESCE((CASE WHEN :SYANAI = '2' THEN  0" & vbCrLf &
    '                        "                            ELSE " & vbCrLf &
    '                        "                               CASE WHEN m2.notax_flg = '0' THEN ROUND((t1.futai_kin * m4.tax_ritu) /100 ,0)" & vbCrLf &
    '                        "                                    ELSE 0 " & vbCrLf &
    '                        "                               END " & vbCrLf &
    '                        "     END),0)  as tax_kin " & vbCrLf &
    '                        "FROM " & vbCrLf &
    '                        "    friyo_meisai_tbl t1 " & vbCrLf &
    '                        "LEFT JOIN futai_mst m1 " & vbCrLf &
    '                        "    ON t1.futai_bunrui_cd = m1.bunrui_cd " & vbCrLf &
    '                        "    AND t1.futai_cd = m1.futai_cd " & vbCrLf &
    '                        "    AND m1.shisetu_kbn = :ShisetuKbn " & vbCrLf &
    '                        "    AND :Riyobi BETWEEN m1.kikan_from AND m1.kikan_to " & vbCrLf &
    '                        "    AND m1.sts = '0' " & vbCrLf &
    '                        "LEFT JOIN fbunrui_mst m2 " & vbCrLf &
    '                        "    ON t1.futai_bunrui_cd = m2.bunrui_cd " & vbCrLf &
    '                        "    AND m2.shisetu_kbn = :ShisetuKbn " & vbCrLf &
    '                        "    AND :Riyobi BETWEEN m2.kikan_from AND m2.kikan_to " & vbCrLf &
    '                        "    AND m2.sts = '0' " & vbCrLf &
    '                        "LEFT JOIN kamoku_mst m3 " & vbCrLf &
    '                        "    ON m2.kamoku_cd = m3.kamoku_cd " & vbCrLf &
    '                        "    AND m2.saimoku_cd = m3.saimoku_cd " & vbCrLf &
    '                        "    AND m2.uchi_cd = m3.uchi_cd " & vbCrLf &
    '                        "    AND m2.shosai_cd = m3.shosai_cd " & vbCrLf &
    '                        "    AND m3.sts = '0' " & vbCrLf &
    '                        "LEFT JOIN TAX_MST m4 " & vbCrLf &
    '                        "    ON  :Riyobi BETWEEN m4.TAXS_DT AND m4.TAXE_DT" & vbCrLf &
    '                        "WHERE " & vbCrLf &
    '                        "    t1.yoyaku_no = :YoyakuNo " & vbCrLf &
    '                        "    AND t1.yoyaku_dt = :Riyobi " & vbCrLf &
    '                        "ORDER BY t1.add_dt "

    ' --- 2019/07/24 軽減税率対応 End E.Okuda@Compass ---

    'Private strEX05S097 As String = _
    '                        "SELECT " & vbCrLf & _
    '                        "    '' as yoyaku_no, " & vbCrLf & _
    '                        "    :Riyobi as yoyaku_dt,  " & vbCrLf & _
    '                        "    SUM(tanka) as total_fuzoku_kin,  " & vbCrLf & _
    '                        "    ARRAY_TO_STRING(ARRAY_AGG(futai_nm), '／') as fuzoku_nm " & vbCrLf & _
    '                        "FROM " & vbCrLf & _
    '                        "    futai_mst " & vbCrLf & _
    '                        "WHERE " & vbCrLf & _
    '                        "    def_flg = '1' " & vbCrLf & _
    '                        "AND sts = '0' " & vbCrLf & _
    '                        "AND :Riyobi BETWEEN kikan_from AND kikan_to " & vbCrLf & _
    '                        "AND shisetu_kbn = :ShisetuKbn " & vbCrLf & _
    '                        "group by " & vbCrLf & _
    '                        "    def_flg, " & vbCrLf & _
    '                        "    sts "
    Private strEX05S097 As String =
                            "SELECT " & vbCrLf &
                            "    '' as yoyaku_no, " & vbCrLf &
                            "    :Riyobi as yoyaku_dt,  " & vbCrLf &
                            "    SUM(m1.tanka) as total_fuzoku_kin,  " & vbCrLf &
                            "    ARRAY_TO_STRING(ARRAY_AGG(m1.futai_nm), '／') as fuzoku_nm " & vbCrLf &
                            "   ,CASE WHEN :SYANAI = '2' THEN  0" & vbCrLf &
                            "                            ELSE  " & vbCrLf &
                            "                              SUM( CASE WHEN m2.NOTAX_FLG = '0' THEN ROUND((m1.TANKA * m3.tax_ritu) /100 ,0)" & vbCrLf &
                            "                                                                ELSE 0 " & vbCrLf &
                            "                                    END ) " & vbCrLf &
                            "    END  AS TAX_KIN " & vbCrLf &
                            "FROM " & vbCrLf &
                            "    futai_mst m1" & vbCrLf &
                                       "           LEFT OUTER JOIN  FBUNRUI_MST m2 " & vbCrLf &
                                       "                       ON   m1.BUNRUI_CD = m2.BUNRUI_CD " & vbCrLf &
                                       "                      AND   m1.SHISETU_KBN = m2.SHISETU_KBN " & vbCrLf &
                                       "                      AND   :Riyobi BETWEEN m2.KIKAN_FROM AND m2.KIKAN_TO " & vbCrLf &
                                       "           LEFT OUTER JOIN  TAX_MST m3 " & vbCrLf &
                                       "                       ON   :Riyobi BETWEEN m3.TAXS_DT AND m3.TAXE_DT " & vbCrLf &
                            "WHERE " & vbCrLf &
                            "    m1.def_flg = '1' " & vbCrLf &
                            "AND m1.sts = '0' " & vbCrLf &
                            "AND :Riyobi BETWEEN m1.kikan_from AND m1.kikan_to " & vbCrLf &
                            "AND m1.shisetu_kbn = :ShisetuKbn " & vbCrLf &
                            "group by " & vbCrLf &
                            "    m1.def_flg, " & vbCrLf &
                            "    m1.sts "

    '付帯設備明細初期値の取得
    ' --- 2019/07/24 軽減税率対応 Start E.Okuda@Compass ---
    Private strEX05S096 As String =
                            "SELECT " & vbCrLf &
                            "    '' as yoyaku_no,  " & vbCrLf &
                            "    :Riyobi as yoyaku_dt,  " & vbCrLf &
                            "    m1.bunrui_cd as futai_bunrui_cd, " & vbCrLf &
                            "    m1.futai_cd, " & vbCrLf &
                            "    m1.tanka as futai_tanka, " & vbCrLf &
                            "    1 as futai_su, " & vbCrLf &
                            "    CAST(m1.tanka AS INT8) as futai_shokei, " & vbCrLf &
                            "    0 as futai_chosei, " & vbCrLf &
                            "    CAST(m1.tanka AS INT8) as futai_kin, " & vbCrLf &
                            "    '' as futai_biko, " & vbCrLf &
                            "    m1.futai_nm, " & vbCrLf &
                            "    m1.tani as futai_tani, " & vbCrLf &
                            "    m2.bunrui_nm as futai_bunrui_nm, " & vbCrLf &
                            "    m2.shukei_grp, " & vbCrLf &
                            "    m2.kamoku_cd, " & vbCrLf &
                            "    m2.saimoku_cd, " & vbCrLf &
                            "    m2.uchi_cd, " & vbCrLf &
                            "    m2.shosai_cd, " & vbCrLf &
                            "    m2.karikamoku_cd, " & vbCrLf &
                            "    m2.kari_saimoku_cd, " & vbCrLf &
                            "    m2.kari_uchi_cd, " & vbCrLf &
                            "    m2.kari_shosai_cd, " & vbCrLf &
                            "    m2.notax_flg, " & vbCrLf &
                            "    (CASE WHEN m2.zeiritsu IS NULL THEN m4.tax_ritu ELSE m2.zeiritsu END) as zeiritsu, " & vbCrLf &
                            "    m3.kamoku_nm, " & vbCrLf &
                            "    m3.saimoku_nm, " & vbCrLf &
                            "    m3.uchi_nm, " & vbCrLf &
                            "    m3.shosai_nm " & vbCrLf &
                            "   ,CASE WHEN :SYANAI = '2' THEN  0" & vbCrLf &
                            "                            ELSE " & vbCrLf &
                            "                                      CASE WHEN m2.NOTAX_FLG = '0' THEN ROUND((m1.TANKA *(CASE WHEN m2.zeiritsu IS NULL THEN m4.tax_ritu ELSE m2.zeiritsu END) ) /100 ,0)" & vbCrLf &
                            "                                      ELSE 0 " & vbCrLf &
                            "                               END " & vbCrLf &
                            "    END  AS TAX_KIN " & vbCrLf &
                            "FROM " & vbCrLf &
                            "    futai_mst m1 " & vbCrLf &
                            "LEFT JOIN fbunrui_mst m2 " & vbCrLf &
                            "    ON m1.bunrui_cd = m2.bunrui_cd " & vbCrLf &
                            "    AND m2.shisetu_kbn = :ShisetuKbn " & vbCrLf &
                            "    AND :Riyobi BETWEEN m2.kikan_from AND m2.kikan_to " & vbCrLf &
                            "    AND m2.sts = '0' " & vbCrLf &
                            "LEFT JOIN kamoku_mst m3 " & vbCrLf &
                            "    ON m2.kamoku_cd = m3.kamoku_cd " & vbCrLf &
                            "    AND m2.saimoku_cd = m3.saimoku_cd " & vbCrLf &
                            "    AND m2.uchi_cd = m3.uchi_cd " & vbCrLf &
                            "    AND m2.shosai_cd = m3.shosai_cd " & vbCrLf &
                            "    AND m3.sts = '0' " & vbCrLf &
                            "LEFT OUTER JOIN  TAX_MST m4 " & vbCrLf &
                            "    ON  :Riyobi BETWEEN m4.TAXS_DT AND m4.TAXE_DT " & vbCrLf &
                            "WHERE " & vbCrLf &
                            "    m1.def_flg = '1' " & vbCrLf &
                            "AND m1.sts = '0' " & vbCrLf &
                            "AND :Riyobi BETWEEN m1.kikan_from AND m1.kikan_to " & vbCrLf &
                            "AND m1.shisetu_kbn = :ShisetuKbn " & vbCrLf &
                            "ORDER BY " & vbCrLf &
                            "    m1.sort "
    'Private strEX05S096 As String =
    '                        "SELECT " & vbCrLf &
    '                        "    '' as yoyaku_no,  " & vbCrLf &
    '                        "    :Riyobi as yoyaku_dt,  " & vbCrLf &
    '                        "    m1.bunrui_cd as futai_bunrui_cd, " & vbCrLf &
    '                        "    m1.futai_cd, " & vbCrLf &
    '                        "    m1.tanka as futai_tanka, " & vbCrLf &
    '                        "    1 as futai_su, " & vbCrLf &
    '                        "    CAST(m1.tanka AS INT8) as futai_shokei, " & vbCrLf &
    '                        "    0 as futai_chosei, " & vbCrLf &
    '                        "    CAST(m1.tanka AS INT8) as futai_kin, " & vbCrLf &
    '                        "    '' as futai_biko, " & vbCrLf &
    '                        "    m1.futai_nm, " & vbCrLf &
    '                        "    m1.tani as futai_tani, " & vbCrLf &
    '                        "    m2.bunrui_nm as futai_bunrui_nm, " & vbCrLf &
    '                        "    m2.shukei_grp, " & vbCrLf &
    '                        "    m2.kamoku_cd, " & vbCrLf &
    '                        "    m2.saimoku_cd, " & vbCrLf &
    '                        "    m2.uchi_cd, " & vbCrLf &
    '                        "    m2.shosai_cd, " & vbCrLf &
    '                        "    m2.karikamoku_cd, " & vbCrLf &
    '                        "    m2.kari_saimoku_cd, " & vbCrLf &
    '                        "    m2.kari_uchi_cd, " & vbCrLf &
    '                        "    m2.kari_shosai_cd, " & vbCrLf &
    '                        "    m2.notax_flg, " & vbCrLf &
    '                        "    m3.kamoku_nm, " & vbCrLf &
    '                        "    m3.saimoku_nm, " & vbCrLf &
    '                        "    m3.uchi_nm, " & vbCrLf &
    '                        "    m3.shosai_nm " & vbCrLf &
    '                        "   ,CASE WHEN :SYANAI = '2' THEN  0" & vbCrLf &
    '                        "                            ELSE " & vbCrLf &
    '                        "                                      CASE WHEN m2.NOTAX_FLG = '0' THEN ROUND((m1.TANKA * m4.tax_ritu) /100 ,0)" & vbCrLf &
    '                        "                                      ELSE 0 " & vbCrLf &
    '                        "                               END " & vbCrLf &
    '                        "    END  AS TAX_KIN " & vbCrLf &
    '                        "FROM " & vbCrLf &
    '                        "    futai_mst m1 " & vbCrLf &
    '                        "LEFT JOIN fbunrui_mst m2 " & vbCrLf &
    '                        "    ON m1.bunrui_cd = m2.bunrui_cd " & vbCrLf &
    '                        "    AND m2.shisetu_kbn = :ShisetuKbn " & vbCrLf &
    '                        "    AND :Riyobi BETWEEN m2.kikan_from AND m2.kikan_to " & vbCrLf &
    '                        "    AND m2.sts = '0' " & vbCrLf &
    '                        "LEFT JOIN kamoku_mst m3 " & vbCrLf &
    '                        "    ON m2.kamoku_cd = m3.kamoku_cd " & vbCrLf &
    '                        "    AND m2.saimoku_cd = m3.saimoku_cd " & vbCrLf &
    '                        "    AND m2.uchi_cd = m3.uchi_cd " & vbCrLf &
    '                        "    AND m2.shosai_cd = m3.shosai_cd " & vbCrLf &
    '                        "    AND m3.sts = '0' " & vbCrLf &
    '                        "LEFT OUTER JOIN  TAX_MST m4 " & vbCrLf &
    '                        "    ON  :Riyobi BETWEEN m4.TAXS_DT AND m4.TAXE_DT " & vbCrLf &
    '                        "WHERE " & vbCrLf &
    '                        "    m1.def_flg = '1' " & vbCrLf &
    '                        "AND m1.sts = '0' " & vbCrLf &
    '                        "AND :Riyobi BETWEEN m1.kikan_from AND m1.kikan_to " & vbCrLf &
    '                        "AND m1.shisetu_kbn = :ShisetuKbn " & vbCrLf &
    '                        "ORDER BY " & vbCrLf &
    '                        "    m1.sort "
    ' --- 2019/07/24 軽減税率対応 End E.Okuda@Compass ---

    '消費税の取得
    Private strEX05S095 As String = _
                            "SELECT " & vbCrLf & _
                            "    tax_ritu " & vbCrLf & _
                            "FROM " & vbCrLf & _
                            "    tax_mst " & vbCrLf & _
                            "WHERE " & vbCrLf & _
                            " :Riyobi BETWEEN taxs_dt AND taxe_dt "

    '承認依頼
    Private strEX05S006 As String = _
                            "SELECT " & vbCrLf & _
                            "    t1.yoyaku_no, " & vbCrLf & _
                            "    t1.seq, " & vbCrLf & _
                            "    to_char(t1.irai_dt, 'yyyy/mm/dd hh24:mi') as irai_dt, " & vbCrLf & _
                            "    t1.com, " & vbCrLf & _
                            "    t2.user_nm as user_nm " & vbCrLf & _
                            "FROM " & vbCrLf & _
                            "    irai_rireki_tbl t1 " & vbCrLf & _
                            "    LEFT JOIN " & vbCrLf & _
                            "        user_mst t2 " & vbCrLf & _
                            "        ON  t1.irai_usercd = t2.user_cd " & vbCrLf & _
                            "WHERE yoyaku_no = :YoyakuNo ORDER BY t1.seq "

    '承認記録
    Private strEX05S007 As String = _
                            "SELECT " & vbCrLf & _
                            "    t1.yoyaku_no, " & vbCrLf & _
                            "    t1.seq, " & vbCrLf & _
                            "    to_char(t1.check_dt, 'yyyy/mm/dd hh24:mi') as check_dt, " & vbCrLf & _
                            "    t1.check_sts, " & vbCrLf & _
                            "    t1.com, " & vbCrLf & _
                            "    t2.user_nm as user_nm " & vbCrLf & _
                            "FROM " & vbCrLf & _
                            "    check_rireki_tbl t1 " & vbCrLf & _
                            "    LEFT JOIN " & vbCrLf & _
                            "        user_mst t2 " & vbCrLf & _
                            "        ON  t1.check_usercd = t2.user_cd " & vbCrLf & _
                            "WHERE t1.yoyaku_no = :YoyakuNo ORDER BY t1.seq "

    'メール送付先
    ' 2015.11.20 UPD START↓ h.hagiwara
    'Private strEX05S099 As String = _
    '                        "SELECT " & vbCrLf & _
    '                        "    mail " & vbCrLf & _
    '                        "FROM " & vbCrLf & _
    '                        "    user_mst " & vbCrLf & _
    '                        "WHERE sts = '0' "
    Private strEX05S099 As String = _
                            "SELECT " & vbCrLf & _
                            "    mail " & vbCrLf & _
                            "FROM " & vbCrLf & _
                            "    user_mst " & vbCrLf & _
                            "WHERE sts = '0' " & vbCrLf & _
                            " AND  mail is not null " & vbCrLf & _
                            " AND  user_cd <> '000000' "
    ' 2015.11.20 UPD END↑ h.hagiwara
    ' 2016.01.27 ADD START↓ h.hagiwara
    Private strEX05S0991 As String = " UNION " & vbCrLf & _
                                     "SELECT " & vbCrLf & _
                                     "     MAIL " & vbCrLf & _
                                     "FROM " & vbCrLf & _
                                     "    USER_MST " & vbCrLf & _
                                     "WHERE " & vbCrLf & _
                                     "      STS = '0' " & vbCrLf & _
                                     " AND  mail is not null " & vbCrLf & _
                                     " AND  user_cd = ( SELECT " & vbCrLf & _
                                     "                     IRAI_USERCD " & vbCrLf & _
                                     "                  FROM " & vbCrLf & _
                                     "                     IRAI_RIREKI_TBL " & vbCrLf & _
                                     "                  WHERE " & vbCrLf & _
                                     "                      YOYAKU_NO = :YOYAKU_NO " & vbCrLf & _
                                     "                  AND SEQ = ( SELECT MAX(SEQ) FROM IRAI_RIREKI_TBL " & vbCrLf & _
                                     "                                              WHERE YOYAKU_NO = :YOYAKU_NO2 )" & vbCrLf & _
                                     "                 )"
    ' 2016.01.27 ADD END↑ h.hagiwara

    '利用料(初回)取得
    Private strEX38S001 As String = _
                            "SELECT " & vbCrLf & _
                            "    null as yoyaku_no, " & vbCrLf & _
                            "    null as seq, " & vbCrLf & _
                            "    null as seikyu_irai_no, " & vbCrLf & _
                            "    null as seikyu_irai_seq, " & vbCrLf & _
                            "    '1' as seikyu_naiyo, " & vbCrLf & _
                            "    '' as riyo_ym, " & vbCrLf & _
                            "    t1.shukei_grp, " & vbCrLf & _
                            "    t2.kamoku_nm, " & vbCrLf & _
                            "    t2.saimoku_nm, " & vbCrLf & _
                            "    t2.uchi_nm, " & vbCrLf & _
                            "    t2.shosai_nm, " & vbCrLf & _
                            "    CAST(0 AS INT8) as keijo_kin, " & vbCrLf & _
                            "    CAST(0 AS INT8) as tax_kin, " & vbCrLf & _
                            "    '' as tekiyo1, " & vbCrLf & _
                            "    '' as tekiyo2, " & vbCrLf & _
                            "    '' as event_nm, " & vbCrLf & _
                            "    '' as content_uchi_nm, " & vbCrLf & _
                            "    t2.kamoku_cd, " & vbCrLf & _
                            "    COALESCE(t2.saimoku_cd,'') as saimoku_cd , " & vbCrLf & _
                            "    COALESCE(t2.uchi_cd,'') as uchi_cd , " & vbCrLf & _
                            "    COALESCE(t2.shosai_cd,'') as shosai_cd , " & vbCrLf & _
                            "    '' as content_cd, " & vbCrLf & _
                            "    '' as content_uchi_cd, " & vbCrLf & _
                            "    t1.shukei_grp, " & vbCrLf & _
                            "    t2.karikamoku_cd, " & vbCrLf & _
                            "    t2.kari_saimoku_cd, " & vbCrLf & _
                            "    t2.kari_uchi_cd, " & vbCrLf & _
                            "    t2.kari_shosai_cd " & vbCrLf & _
                            "FROM " & vbCrLf & _
                            "    fbunrui_mst t1 " & vbCrLf & _
                            "INNER JOIN kamoku_mst t2 " & vbCrLf & _
                            "ON        t1.kamoku_cd = t2.kamoku_cd" & vbCrLf & _
                            "     AND  t1.saimoku_cd = t2.saimoku_cd" & vbCrLf & _
                            "     AND  t1.uchi_cd = t2.uchi_cd" & vbCrLf & _
                            "     AND  t1.shosai_cd = t2.shosai_cd" & vbCrLf & _
                            "WHERE " & vbCrLf & _
                            "        t2.riyo_kamoku_flg = '1' "

    '付帯(初回)取得
    Private strEX38S002 As String = _
                            "SELECT " & vbCrLf & _
                            "    null as yoyaku_no, " & vbCrLf & _
                            "    null as seq, " & vbCrLf & _
                            "    null as seikyu_irai_no, " & vbCrLf & _
                            "    null as seikyu_irai_seq, " & vbCrLf & _
                            "    '2' as seikyu_naiyo, " & vbCrLf & _
                            "    '' as riyo_ym, " & vbCrLf & _
                            "    t1.shukei_grp, " & vbCrLf & _
                            "    t2.kamoku_nm, " & vbCrLf & _
                            "    t2.saimoku_nm, " & vbCrLf & _
                            "    t2.uchi_nm, " & vbCrLf & _
                            "    t2.shosai_nm, " & vbCrLf & _
                            "    CAST(0 AS INT8) as keijo_kin, " & vbCrLf & _
                            "    CAST(0 AS INT8) as tax_kin, " & vbCrLf & _
                            "    '' as tekiyo1, " & vbCrLf & _
                            "    '' as tekiyo2, " & vbCrLf & _
                            "    '' as event_nm, " & vbCrLf & _
                            "    '' as content_uchi_nm, " & vbCrLf & _
                            "    t2.kamoku_cd, " & vbCrLf & _
                            "    COALESCE(t2.saimoku_cd,'') as saimoku_cd , " & vbCrLf & _
                            "    COALESCE(t2.uchi_cd,'') as uchi_cd , " & vbCrLf & _
                            "    COALESCE(t2.shosai_cd,'') as shosai_cd , " & vbCrLf & _
                            "    '' as content_cd, " & vbCrLf & _
                            "    '' as content_uchi_cd, " & vbCrLf & _
                            "    t1.shukei_grp, " & vbCrLf & _
                            "    COALESCE(t2.karikamoku_cd,'') as karikamoku_cd , " & vbCrLf & _
                            "    COALESCE(t2.kari_saimoku_cd,'') as kari_saimoku_cd , " & vbCrLf & _
                            "    COALESCE(t2.kari_uchi_cd,'') as kari_uchi_cd , " & vbCrLf & _
                            "    COALESCE(t2.kari_shosai_cd,'') as kari_shosai_cd  " & vbCrLf & _
                            "FROM " & vbCrLf & _
                            "    fbunrui_mst t1 " & vbCrLf & _
                            "INNER JOIN kamoku_mst t2 " & vbCrLf & _
                            "ON        t1.kamoku_cd = t2.kamoku_cd" & vbCrLf & _
                            "     AND  t1.saimoku_cd = t2.saimoku_cd" & vbCrLf & _
                            "     AND  t1.uchi_cd = t2.uchi_cd" & vbCrLf & _
                            "     AND  t1.shosai_cd = t2.shosai_cd" & vbCrLf & _
                            "WHERE " & vbCrLf & _
                            "     :Riyobi BETWEEN t1.kikan_from AND t1.kikan_to " & vbCrLf & _
                            "AND  t1.shisetu_kbn = :ShisetuKbn " & vbCrLf & _
                            "AND  t1.sts = '0' " & vbCrLf & _
                            " ORDER BY " & vbCrLf & _
                            "   riyo_ym ASC , t2.kamoku_nm ASC "

    '利用料(初回以降)取得
    Private strEX38S004 As String = _
                            "SELECT " & vbCrLf & _
                            "    t1.yoyaku_no, " & vbCrLf & _
                            "    t1.seq, " & vbCrLf & _
                            "    t1.seikyu_irai_no, " & vbCrLf & _
                            "    t1.seikyu_irai_seq, " & vbCrLf & _
                            "    t1.seikyu_naiyo, " & vbCrLf & _
                            "    t1.riyo_ym, " & vbCrLf & _
                            "    t2.kamoku_nm, " & vbCrLf & _
                            "    t2.saimoku_nm, " & vbCrLf & _
                            "    t2.uchi_nm, " & vbCrLf & _
                            "    t2.shosai_nm, " & vbCrLf & _
                            "    t1.keijo_kin, " & vbCrLf & _
                            "    COALESCE(t1.tax_kin,0) tax_kin ," & vbCrLf & _
                            "    t1.tekiyo1, " & vbCrLf & _
                            "    t1.tekiyo2, " & vbCrLf & _
                            "    t3.event_nm, " & vbCrLf & _
                            "    t4.content_uchi_nm, " & vbCrLf & _
                            "    t1.kamoku_cd, " & vbCrLf & _
                            "    t1.saimoku_cd, " & vbCrLf & _
                            "    t1.uchi_cd, " & vbCrLf & _
                            "    t1.shosai_cd, " & vbCrLf & _
                            "    t1.content_cd, " & vbCrLf & _
                            "    t1.content_uchi_cd, " & vbCrLf & _
                            "    t1.shukei_grp, " & vbCrLf & _
                            "    t2.karikamoku_cd, " & vbCrLf & _
                            "    t2.kari_saimoku_cd, " & vbCrLf & _
                            "    t2.kari_uchi_cd, " & vbCrLf & _
                            "    t2.kari_shosai_cd " & vbCrLf & _
                            "FROM " & vbCrLf & _
                            "    project_tbl t1 " & vbCrLf & _
                            "    LEFT JOIN " & vbCrLf & _
                            "        kamoku_mst t2 " & vbCrLf & _
                            "    ON  t1.kamoku_cd = t2.kamoku_cd " & vbCrLf & _
                            "    AND t1.saimoku_cd = t2.saimoku_cd " & vbCrLf & _
                            "    AND t1.uchi_cd = t2.uchi_cd " & vbCrLf & _
                            "    AND t1.shosai_cd = t2.shosai_cd " & vbCrLf & _
                            "    LEFT JOIN " & vbCrLf & _
                            "        content_mst t3 " & vbCrLf & _
                            "    ON  t1.content_cd = t3.event_cd " & vbCrLf & _
                            "    LEFT JOIN " & vbCrLf & _
                            "        content_uchi_mst t4 " & vbCrLf & _
                            "    ON  t1.content_cd = t4.content_cd " & vbCrLf & _
                            "    AND t1.content_uchi_cd = t4.content_uchi_cd " & vbCrLf & _
                            "WHERE " & vbCrLf & _
                            "    ( " & vbCrLf & _
                            "        t1.seikyu_naiyo = '1' " & vbCrLf & _
                            "    OR  t1.seikyu_naiyo = '3' " & vbCrLf & _
                            "    ) " & vbCrLf & _
                            "AND t1.yoyaku_no = :YoyakuNo " & vbCrLf & _
                            "AND t1.seikyu_irai_no = :SeikyuIraiNo " & vbCrLf & _
                            " ORDER BY " & vbCrLf & _
                            "   t1.riyo_ym ASC , t2.kamoku_nm ASC "

    '付帯設備(初回以降)取得
    ' --- 2019/08/08 軽減税率対応 Start E.Okuda@Compass ---
    Private strEX38S005 As String =
                            "SELECT " & vbCrLf &
                            "    t1.yoyaku_no, " & vbCrLf &
                            "    t1.seq, " & vbCrLf &
                            "    t1.seikyu_irai_no, " & vbCrLf &
                            "    t1.seikyu_irai_seq, " & vbCrLf &
                            "    t1.seikyu_naiyo, " & vbCrLf &
                            "    t1.riyo_ym, " & vbCrLf &
                            "    t1.shukei_grp, " & vbCrLf &
                            "    t2.kamoku_nm, " & vbCrLf &
                            "    t2.saimoku_nm, " & vbCrLf &
                            "    t2.uchi_nm, " & vbCrLf &
                            "    t2.shosai_nm, " & vbCrLf &
                            "    t1.keijo_kin, " & vbCrLf &
                            "    COALESCE(t1.tax_kin,0) tax_kin ," & vbCrLf &
                            "    t1.tekiyo1, " & vbCrLf &
                            "    t1.tekiyo2, " & vbCrLf &
                            "    t3.event_nm, " & vbCrLf &
                            "    t4.content_uchi_nm, " & vbCrLf &
                            "    t1.kamoku_cd, " & vbCrLf &
                            "    t1.saimoku_cd, " & vbCrLf &
                            "    t1.uchi_cd, " & vbCrLf &
                            "    t1.shosai_cd, " & vbCrLf &
                            "    t1.content_cd, " & vbCrLf &
                            "    t1.content_uchi_cd, " & vbCrLf &
                            "    t1.shukei_grp, " & vbCrLf &
                            "    t2.karikamoku_cd, " & vbCrLf &
                            "    t2.kari_saimoku_cd, " & vbCrLf &
                            "    t2.kari_uchi_cd, " & vbCrLf &
                            "    t2.kari_shosai_cd, " & vbCrLf &
                            "    '' as zeiritsu, " & vbCrLf &
                            "    '' as notax_flg " & vbCrLf &
                            "FROM " & vbCrLf &
                            "    project_tbl t1 " & vbCrLf &
                            "    LEFT JOIN " & vbCrLf &
                            "        kamoku_mst t2 " & vbCrLf &
                            "    ON  t1.kamoku_cd = t2.kamoku_cd " & vbCrLf &
                            "    AND t1.saimoku_cd = t2.saimoku_cd " & vbCrLf &
                            "    AND t1.uchi_cd = t2.uchi_cd " & vbCrLf &
                            "    AND t1.shosai_cd = t2.shosai_cd " & vbCrLf &
                            "    LEFT JOIN " & vbCrLf &
                            "        content_mst t3 " & vbCrLf &
                            "    ON  t1.content_cd = t3.event_cd " & vbCrLf &
                            "    LEFT JOIN " & vbCrLf &
                            "        content_uchi_mst t4 " & vbCrLf &
                            "    ON  t1.content_cd = t4.content_cd " & vbCrLf &
                            "    AND t1.content_uchi_cd = t4.content_uchi_cd " & vbCrLf &
                            "WHERE " & vbCrLf &
                            "    ( " & vbCrLf &
                            "        t1.seikyu_naiyo = '2' " & vbCrLf &
                            "    OR  t1.seikyu_naiyo = '3' " & vbCrLf &
                            "    ) " & vbCrLf &
                            "AND t1.yoyaku_no = :YoyakuNo " & vbCrLf &
                            "AND t1.seikyu_irai_no = :SeikyuIraiNo " & vbCrLf &
                            " ORDER BY " & vbCrLf &
                            "   t1.riyo_ym ASC , t2.kamoku_nm ASC "
    'Private strEX38S005 As String =
    '                        "SELECT " & vbCrLf &
    '                        "    t1.yoyaku_no, " & vbCrLf &
    '                        "    t1.seq, " & vbCrLf &
    '                        "    t1.seikyu_irai_no, " & vbCrLf &
    '                        "    t1.seikyu_irai_seq, " & vbCrLf &
    '                        "    t1.seikyu_naiyo, " & vbCrLf &
    '                        "    t1.riyo_ym, " & vbCrLf &
    '                        "    t1.shukei_grp, " & vbCrLf &
    '                        "    t2.kamoku_nm, " & vbCrLf &
    '                        "    t2.saimoku_nm, " & vbCrLf &
    '                        "    t2.uchi_nm, " & vbCrLf &
    '                        "    t2.shosai_nm, " & vbCrLf &
    '                        "    t1.keijo_kin, " & vbCrLf &
    '                        "    COALESCE(t1.tax_kin,0) tax_kin ," & vbCrLf &
    '                        "    t1.tekiyo1, " & vbCrLf &
    '                        "    t1.tekiyo2, " & vbCrLf &
    '                        "    t3.event_nm, " & vbCrLf &
    '                        "    t4.content_uchi_nm, " & vbCrLf &
    '                        "    t1.kamoku_cd, " & vbCrLf &
    '                        "    t1.saimoku_cd, " & vbCrLf &
    '                        "    t1.uchi_cd, " & vbCrLf &
    '                        "    t1.shosai_cd, " & vbCrLf &
    '                        "    t1.content_cd, " & vbCrLf &
    '                        "    t1.content_uchi_cd, " & vbCrLf &
    '                        "    t1.shukei_grp, " & vbCrLf &
    '                        "    t2.karikamoku_cd, " & vbCrLf &
    '                        "    t2.kari_saimoku_cd, " & vbCrLf &
    '                        "    t2.kari_uchi_cd, " & vbCrLf &
    '                        "    t2.kari_shosai_cd " & vbCrLf &
    '                        "FROM " & vbCrLf &
    '                        "    project_tbl t1 " & vbCrLf &
    '                        "    LEFT JOIN " & vbCrLf &
    '                        "        kamoku_mst t2 " & vbCrLf &
    '                        "    ON  t1.kamoku_cd = t2.kamoku_cd " & vbCrLf &
    '                        "    AND t1.saimoku_cd = t2.saimoku_cd " & vbCrLf &
    '                        "    AND t1.uchi_cd = t2.uchi_cd " & vbCrLf &
    '                        "    AND t1.shosai_cd = t2.shosai_cd " & vbCrLf &
    '                        "    LEFT JOIN " & vbCrLf &
    '                        "        content_mst t3 " & vbCrLf &
    '                        "    ON  t1.content_cd = t3.event_cd " & vbCrLf &
    '                        "    LEFT JOIN " & vbCrLf &
    '                        "        content_uchi_mst t4 " & vbCrLf &
    '                        "    ON  t1.content_cd = t4.content_cd " & vbCrLf &
    '                        "    AND t1.content_uchi_cd = t4.content_uchi_cd " & vbCrLf &
    '                        "WHERE " & vbCrLf &
    '                        "    ( " & vbCrLf &
    '                        "        t1.seikyu_naiyo = '2' " & vbCrLf &
    '                        "    OR  t1.seikyu_naiyo = '3' " & vbCrLf &
    '                        "    ) " & vbCrLf &
    '                        "AND t1.yoyaku_no = :YoyakuNo " & vbCrLf &
    '                        "AND t1.seikyu_irai_no = :SeikyuIraiNo " & vbCrLf &
    '                        " ORDER BY " & vbCrLf &
    '                        "   t1.riyo_ym ASC , t2.kamoku_nm ASC "
    ' --- 2019/08/08 軽減税率対応 End E.Okuda@Compass ---

    'SQL文(入金情報)
    Private strEX41S001 As String = _
                            "SELECT " & vbCrLf & _
                            "    yoyaku_no, " & vbCrLf & _
                            "    aite_cd, " & vbCrLf & _
                            "    aite_nm, " & vbCrLf & _
                            "    nyukin_yotei_dt, " & vbCrLf & _
                            "    nyukin_dt, " & vbCrLf & _
                            "    seikyu_kin, " & vbCrLf & _
                            "    seikyu_dt, " & vbCrLf & _
                            "    input_dt, " & vbCrLf & _
                            "    sekikyu_no, " & vbCrLf & _
                            "    seikyu_irai_no, " & vbCrLf & _
                            "    seikyu_kin, " & vbCrLf & _
                            "    nyukin_link_no, " & vbCrLf & _
                            "    :LineNo as line_no " & vbCrLf & _
                            "FROM " & vbCrLf & _
                            "    exas_nyukin_tbl " & vbCrLf & _
                            "WHERE 1 = 1 "

    'コンテンツ情報
    Private strEX41S006 As String = _
                           "SELECT " & vbCrLf & _
                           "    m1.event_cd, " & vbCrLf & _
                           "    m1.event_nm, " & vbCrLf & _
                           "    m2.content_uchi_cd, " & vbCrLf & _
                           "    m2.content_uchi_nm " & vbCrLf & _
                           "FROM CONTENT_MST m1 " & vbCrLf & _
                           "LEFT JOIN CONTENT_UCHI_MST m2 " & vbCrLf & _
                           "    ON m1.event_cd = m2.content_cd " & vbCrLf & _
                           "WHERE (m1.del_flg <> '1' OR m1.del_flg is null ) " & vbCrLf & _
                           "    AND (m2.del_flg <> '1' OR m2.del_flg is null ) " & vbCrLf & _
                           "    AND :Riyobi BETWEEN m1.strat_dt AND m1.end_dt " & vbCrLf & _
                           "    AND :ContentUchiNm = m2.content_uchi_nm "

    '請求情報登録
    Private strEX05D003 As String = _
                            "DELETE FROM BILLPAY_TBL WHERE  yoyaku_no=:YoyakuNo "
    '請求情報登録
    Private strEX05I003 As String = _
                            "INSERT INTO " & vbCrLf & _
                            "    BILLPAY_TBL(yoyaku_no, seq, seikyu_irai_no, seikyu_dt, nyukin_yotei_dt, kakutei_kin, chosei_kin, shokei, " & vbCrLf & _
                            "     tax_kin, seikyu_kin, seikyu_naiyo, aite_cd, nyukin_kbn, seikyu_title1, seikyu_title2, nyukin_dt, nyukin_kin, " & vbCrLf & _
                            "      seikyu_input_flg, seikyu_irai_flg, nyukin_input_flg, nyukin_link_no, add_dt, add_user_cd, up_dt, up_user_cd) " & vbCrLf & _
                            "VALUES " & vbCrLf & _
                            "    ( " & vbCrLf & _
                            "    :YoyakuNo, " & vbCrLf & _
                            "    nextval('BILLPAY_TBL_S'), " & vbCrLf & _
                            "    :SeikyuIraiNo, " & vbCrLf & _
                            "    :SeikyuDt, " & vbCrLf & _
                            "    :NyukinYoteiDt, " & vbCrLf & _
                            "    :KakuteiKin, " & vbCrLf & _
                            "    :ChoseiKin, " & vbCrLf & _
                            "    :Shokei, " & vbCrLf & _
                            "    :TaxKin, " & vbCrLf & _
                            "    :SeikyuKin, " & vbCrLf & _
                            "    :SeikyuNaiyo, " & vbCrLf & _
                            "    :AiteCd, " & vbCrLf & _
                            "    :NyukinKbn, " & vbCrLf & _
                            "    :SeikyuTitle1, " & vbCrLf & _
                            "    :SeikyuTitle2, " & vbCrLf & _
                            "    :NyukinDt, " & vbCrLf & _
                            "    :NyukinKin, " & vbCrLf & _
                            "    :SeikyuInputFlg, " & vbCrLf & _
                            "    :SeikyuIraiFlg, " & vbCrLf & _
                            "    :NyukinInputFlg, " & vbCrLf & _
                            "    :NyukinLinkNo, " & vbCrLf & _
                            "    current_timestamp, " & vbCrLf & _
                            "    :AddUserCd, " & vbCrLf & _
                            "    current_timestamp, " & vbCrLf & _
                            "    :UpUserCd " & vbCrLf & _
                            "    ) RETURNING seq"

    '付帯明細
    Private strEX05D005 As String = _
                            "DELETE FROM FRIYO_TBL WHERE yoyaku_no = :YoyakuNo AND yoyaku_dt not in "

    '付帯登録
    Private strEX05I005 As String = _
                            "INSERT INTO " & vbCrLf & _
                            "    FRIYO_TBL(yoyaku_no, yoyaku_dt, fuzoku_nm, total_fuzoku_kin, add_dt, add_user_cd, up_dt, up_user_cd) " & vbCrLf & _
                            "VALUES " & vbCrLf & _
                            "    ( " & vbCrLf & _
                            "    :YoyakuNo, " & vbCrLf & _
                            "    :Riyobi, " & vbCrLf & _
                            "    :FuzokuNm, " & vbCrLf & _
                            "    :TotalFuzokuKin, " & vbCrLf & _
                            "    current_timestamp, " & vbCrLf & _
                            "    :AddUserCd, " & vbCrLf & _
                            "    current_timestamp, " & vbCrLf & _
                            "    :UpUserCd " & vbCrLf & _
                            "    ) "

    '付帯登録
    Private strEX05U005 As String = _
                            "UPDATE FRIYO_TBL SET " & vbCrLf & _
                            "    fuzoku_nm = :FuzokuNm, " & vbCrLf & _
                            "    total_fuzoku_kin = :TotalFuzokuKin, " & vbCrLf & _
                            "    up_dt = current_timestamp, " & vbCrLf & _
                            "    up_user_cd = :UpUserCd " & vbCrLf & _
                            "WHERE " & vbCrLf & _
                            "    yoyaku_no = :YoyakuNo " & vbCrLf & _
                            "AND yoyaku_dt = :Riyobi "

    '付帯明細
    Private strEX05D005D As String = _
                            "DELETE FROM FRIYO_MEISAI_TBL WHERE yoyaku_no=:YoyakuNo "

    '付帯明細
    Private strEX05I005D As String = _
                            "INSERT INTO " & vbCrLf & _
                            "    FRIYO_MEISAI_TBL(yoyaku_no, yoyaku_dt, futai_bunrui_cd, futai_cd, futai_tanka, futai_su, futai_shokei, futai_chosei, futai_kin, futai_biko, add_dt, add_user_cd, up_dt, up_user_cd) " & vbCrLf & _
                            "VALUES " & vbCrLf & _
                            "    ( " & vbCrLf & _
                            "    :YoyakuNo, " & vbCrLf & _
                            "    :Riyobi, " & vbCrLf & _
                            "    :FutaiBunruiCd, " & vbCrLf & _
                            "    :FutaiCd, " & vbCrLf & _
                            "    :FutaiTanka, " & vbCrLf & _
                            "    :FutaiSu, " & vbCrLf & _
                            "    :FutaiShokei, " & vbCrLf & _
                            "    :FutaiChosei, " & vbCrLf & _
                            "    :FutaiKin, " & vbCrLf & _
                            "    :FutaiBiko, " & vbCrLf & _
                            "    current_timestamp, " & vbCrLf & _
                            "    :AddUserCd, " & vbCrLf & _
                            "    current_timestamp, " & vbCrLf & _
                            "    :UpUserCd " & vbCrLf & _
                            "    ) "

    'プロジェクト設定削除
    ' 2015.12.25 UPD START↓ h.hagiwara
    'Private strEX05D006 As String = _
    '                        "DELETE FROM PROJECT_TBL WHERE yoyaku_no = :YoyakuNo AND seikyu_irai_no = :SeikyuIraiNo AND seq not in "
    Private strEX05D006 As String = _
                            "DELETE FROM PROJECT_TBL WHERE yoyaku_no = :YoyakuNo AND seikyu_irai_no = :SeikyuIraiNo "
    ' 2015.12.25 UPD END↑ h.hagiwara

    'プロジェクト設定登録
    Private strEX05I006 As String = _
                            "INSERT INTO " & vbCrLf & _
                            "    PROJECT_TBL(yoyaku_no, seq, seikyu_irai_no, seikyu_irai_seq, seikyu_naiyo, riyo_ym, keijo_kin, tax_kin, tekiyo1, tekiyo2," & vbCrLf & _
                            "    content_cd, content_uchi_cd, shukei_grp, kamoku_cd, saimoku_cd, uchi_cd, shosai_cd, karikamoku_cd, kari_saimoku_cd," & vbCrLf & _
                            "    kari_uchi_cd, kari_shosai_cd, add_dt, add_user_cd, up_dt, up_user_cd)" & vbCrLf & _
                            "VALUES " & vbCrLf & _
                            "    ( " & vbCrLf & _
                            "    :YoyakuNo, " & vbCrLf & _
                            "    nextval('PROJECT_TBL_S'), " & vbCrLf & _
                            "    :SeikyuIraiNo, " & vbCrLf & _
                            "    :SeikyuIraiSeq, " & vbCrLf & _
                            "    :SeikyuNaiyo, " & vbCrLf & _
                            "    :RiyoYm, " & vbCrLf & _
                            "    :KeijoKin, " & vbCrLf & _
                            "    :TaxKin, " & vbCrLf & _
                            "    :Tekiyo1, " & vbCrLf & _
                            "    :Tekiyo2, " & vbCrLf & _
                            "    :ContentCd, " & vbCrLf & _
                            "    :ContentUchiCd, " & vbCrLf & _
                            "    :ShukeiGrp, " & vbCrLf & _
                            "    :KamokuCd, " & vbCrLf & _
                            "    :SaimokuCd, " & vbCrLf & _
                            "    :UchiCd, " & vbCrLf & _
                            "    :ShosaiCd, " & vbCrLf & _
                            "    :KarikamokuCd, " & vbCrLf & _
                            "    :KariSaimokuCd, " & vbCrLf & _
                            "    :KariUchiCd, " & vbCrLf & _
                            "    :KariShosaiCd, " & vbCrLf & _
                            "    current_timestamp, " & vbCrLf & _
                            "    :AddUserCd, " & vbCrLf & _
                            "    current_timestamp, " & vbCrLf & _
                            "    :UpUserCd " & vbCrLf & _
                            "    ) "


    'プロジェクト設定登録
    Private strEX05U006 As String = _
                            "UPDATE PROJECT_TBL SET " & vbCrLf & _
                            "    seikyu_naiyo = :SeikyuNaiyo, " & vbCrLf & _
                            "    riyo_ym = :RiyoYm, " & vbCrLf & _
                            "    keijo_kin = :KeijoKin, " & vbCrLf & _
                            "    tax_kin = :TaxKin, " & vbCrLf & _
                            "    tekiyo1 = :Tekiyo1, " & vbCrLf & _
                            "    tekiyo2 = :Tekiyo2, " & vbCrLf & _
                            "    content_cd = :ContentCd, " & vbCrLf & _
                            "    content_uchi_cd =:ContentUchiCd, " & vbCrLf & _
                            "    shukei_grp = :ShukeiGrp, " & vbCrLf & _
                            "    kamoku_cd = :KamokuCd, " & vbCrLf & _
                            "    saimoku_cd = :SaimokuCd, " & vbCrLf & _
                            "    uchi_cd = :UchiCd, " & vbCrLf & _
                            "    shosai_cd = :ShosaiCd, " & vbCrLf & _
                            "    karikamoku_cd = :KarikamokuCd, " & vbCrLf & _
                            "    kari_saimoku_cd =:KariSaimokuCd, " & vbCrLf & _
                            "    kari_uchi_cd = :KariUchiCd, " & vbCrLf & _
                            "    kari_shosai_cd = :KariShosaiCd, " & vbCrLf & _
                            "    up_dt = current_timestamp, " & vbCrLf & _
                            "    up_user_cd = :UpUserCd " & vbCrLf & _
                            "WHERE " & vbCrLf & _
                            "    yoyaku_no = :YoyakuNo " & vbCrLf & _
                            "AND seq = :Seq " & vbCrLf & _
                            "AND seikyu_irai_no = :SeikyuIraiNo "

    '依頼履歴
    Private strEX05I010 As String = _
                            "INSERT INTO " & vbCrLf & _
                            "    IRAI_RIREKI_TBL(yoyaku_no, seq, com, irai_dt, irai_usercd)" & vbCrLf & _
                            "VALUES " & vbCrLf & _
                            "    ( " & vbCrLf & _
                            "    :YoyakuNo, " & vbCrLf & _
                            "    nextval('IRAI_RIREKI_TBL_S'), " & vbCrLf & _
                            "    :Com, " & vbCrLf & _
                            "    to_timestamp(:IraiDt,'YYYY/MM/DD HH24:MI'), " & vbCrLf & _
                            "    :IraiUserCd " & vbCrLf & _
                            "    ) RETURNING seq"

    'チェック履歴
    Private strEX05I011 As String = _
                            "INSERT INTO " & vbCrLf & _
                            "    CHECK_RIREKI_TBL(yoyaku_no, seq, check_sts, com, check_dt, check_usercd)" & vbCrLf & _
                            "VALUES " & vbCrLf & _
                            "    ( " & vbCrLf & _
                            "    :YoyakuNo, " & vbCrLf & _
                            "    nextval('CHECK_RIREKI_TBL_S'), " & vbCrLf & _
                            "    :CheckSts, " & vbCrLf & _
                            "    :Com, " & vbCrLf & _
                            "    to_timestamp(:CheckDt,'YYYY/MM/DD HH24:MI'), " & vbCrLf & _
                            "    :CheckUserCd " & vbCrLf & _
                            "    ) RETURNING seq"

    'SQL(予約情報更新)
    '2016.11.4 m.hayabuchi MOD Start 課題No.59
    'Private strEX04U001 As String = _
    '                       "update YOYAKU_TBL " & vbCrLf & _
    '                           "SET " & vbCrLf & _
    '                               "KAKUTEI_DT = :KakuteiDt, " & vbCrLf & _
    '                               "KAKU_USERCD = :KakuUsercd, " & vbCrLf & _
    '                               "YOYAKU_STS = :YoyakuSts, " & vbCrLf & _
    '                               "SHISETU_KBN = :ShisetuKbn, " & vbCrLf & _
    '                               "STUDIO_KBN = :StudioKbn, " & vbCrLf & _
    '                               "SAIJI_NM = :SaijiNm, " & vbCrLf & _
    '                               "SHUTSUEN_NM = :ShutsuenNm, " & vbCrLf & _
    '                               "KASHI_KIND = :KashiKind, " & vbCrLf & _
    '                               "RIYO_TYPE = :RiyoType, " & vbCrLf & _
    '                               "DRINK_FLG = :DrinkFlg, " & vbCrLf & _
    '                               "SAIJI_BUNRUI = :SaijiBunrui, " & vbCrLf & _
    '                               "TEIIN = :Teiin, " & vbCrLf & _
    '                               "ONKYO_OPE_FLG = :OnkyoOpeFlg, " & vbCrLf & _
    '                               "RIYOSHA_CD = :RiyoshaCd, " & vbCrLf & _
    '                               "RIYO_NM = :RiyoNm, " & vbCrLf & _
    '                               "RIYO_KANA = :RiyoKana, " & vbCrLf & _
    '                               "SEKININ_BUSHO_NM = :SekininBushoNm, " & vbCrLf & _
    '                               "SEKININ_NM = :SekininNm, " & vbCrLf & _
    '                               "SEKININ_MAIL = :SekininMail, " & vbCrLf & _
    '                               "DAIHYO_NM = :DaihyoNm, " & vbCrLf & _
    '                               "RIYO_TEL11 = :RiyoTel11, " & vbCrLf & _
    '                               "RIYO_TEL12 = :RiyoTel12, " & vbCrLf & _
    '                               "RIYO_TEL13 = :RiyoTel13, " & vbCrLf & _
    '                               "RIYO_TEL21 = :RiyoTel21, " & vbCrLf & _
    '                               "RIYO_TEL22 = :RiyoTel22, " & vbCrLf & _
    '                               "RIYO_TEL23 = :RiyoTel23, " & vbCrLf & _
    '                               "RIYO_NAISEN = :RiyoNaisen, " & vbCrLf & _
    '                               "RIYO_FAX11 = :RiyoFax11, " & vbCrLf & _
    '                               "RIYO_FAX12 = :RiyoFax12, " & vbCrLf & _
    '                               "RIYO_FAX13 = :RiyoFax13, " & vbCrLf & _
    '                               "RIYO_YUBIN1 = :RiyoYubin1, " & vbCrLf & _
    '                               "RIYO_YUBIN2 = :RiyoYubin2, " & vbCrLf & _
    '                               "RIYO_TODO = :RiyoTodo, " & vbCrLf & _
    '                               "RIYO_SHIKU = :RiyoShiku, " & vbCrLf & _
    '                               "RIYO_BAN = :RiyoBan, " & vbCrLf & _
    '                               "RIYO_BUILD = :RiyoBuild, " & vbCrLf & _
    '                               "RIYO_LVL = :RiyoLvl, " & vbCrLf & _
    '                               "AITE_CD = :AiteCd, " & vbCrLf & _
    '                               "ONKYO_NM = :OnkyoNm, " & vbCrLf & _
    '                               "ONKYO_TANTO_NM = :OnkyoTantoNm, " & vbCrLf & _
    '                               "ONKYO_TEL11 = :OnkyoTel11, " & vbCrLf & _
    '                               "ONKYO_TEL12 = :OnkyoTel12, " & vbCrLf & _
    '                               "ONKYO_TEL13 = :OnkyoTel13, " & vbCrLf & _
    '                               "ONKYO_NAISEN = :OnkyoNaisen, " & vbCrLf & _
    '                               "ONKYO_FAX11 = :OnkyoFax11, " & vbCrLf & _
    '                               "ONKYO_FAX12 = :OnkyoFax12, " & vbCrLf & _
    '                               "ONKYO_FAX13 = :OnkyoFax13, " & vbCrLf & _
    '                               "ONKYO_MAIL = :OnkyoMail, " & vbCrLf & _
    '                               "TOTAL_RIYO_KIN = :TotalRiyoKin, " & vbCrLf & _
    '                               "BIKO = :Biko, " & vbCrLf & _
    '                               "RIYO_COM = :RiyoCom, " & vbCrLf & _
    '                               "TICKET_ENTER_KBN = :TicketEnterKbn, " & vbCrLf & _
    '                               "TICKET_DRINK_KBN = :TicketDrinkKbn, " & vbCrLf & _
    '                               "HP_KEISAI = :HpKeisai, " & vbCrLf & _
    '                               "JOHO_KOKAI_DT = :JohoKokaiDt, " & vbCrLf & _
    '                               "JOHO_KOKAI_TIME = :JohoKokaiTime, " & vbCrLf & _
    '                               "KOKAI_DT = :KokaiDt, " & vbCrLf & _
    '                               "KOKAI_TIME = :KokaiTime, " & vbCrLf & _
    '                               "FINPUT_STS = :FinputSts, " & vbCrLf & _
    '                               "UP_DT = current_timestamp, " & vbCrLf & _
    '                               "UP_USER_CD = :UpUserCd " & vbCrLf & _
    '                           "WHERE " & vbCrLf & _
    '                               "YOYAKU_NO = :YoyakuNo "
    Private strEX04U001 As String = _
                           "update YOYAKU_TBL " & vbCrLf & _
                               "SET " & vbCrLf & _
                                   "KAKUTEI_DT = :KakuteiDt, " & vbCrLf & _
                                   "KAKU_USERCD = :KakuUsercd, " & vbCrLf & _
                                   "YOYAKU_STS = :YoyakuSts, " & vbCrLf & _
                                   "SHISETU_KBN = :ShisetuKbn, " & vbCrLf & _
                                   "STUDIO_KBN = :StudioKbn, " & vbCrLf & _
                                   "SAIJI_NM = :SaijiNm, " & vbCrLf & _
                                   "SHUTSUEN_NM = :ShutsuenNm, " & vbCrLf & _
                                   "KASHI_KIND = :KashiKind, " & vbCrLf & _
                                   "RIYO_TYPE = :RiyoType, " & vbCrLf & _
                                   "DRINK_FLG = :DrinkFlg, " & vbCrLf & _
                                   "SAIJI_BUNRUI = :SaijiBunrui, " & vbCrLf & _
                                   "TEIIN = :Teiin, " & vbCrLf & _
                                   "ONKYO_OPE_FLG = :OnkyoOpeFlg, " & vbCrLf & _
                                   "RIYOSHA_CD = :RiyoshaCd, " & vbCrLf & _
                                   "RIYO_NM = :RiyoNm, " & vbCrLf & _
                                   "RIYO_KANA = :RiyoKana, " & vbCrLf & _
                                   "SEKININ_BUSHO_NM = :SekininBushoNm, " & vbCrLf & _
                                   "SEKININ_NM = :SekininNm, " & vbCrLf & _
                                   "SEKININ_MAIL = :SekininMail, " & vbCrLf & _
                                   "DAIHYO_NM = :DaihyoNm, " & vbCrLf & _
                                   "RIYO_TEL11 = :RiyoTel11, " & vbCrLf & _
                                   "RIYO_TEL12 = :RiyoTel12, " & vbCrLf & _
                                   "RIYO_TEL13 = :RiyoTel13, " & vbCrLf & _
                                   "RIYO_TEL21 = :RiyoTel21, " & vbCrLf & _
                                   "RIYO_NAISEN = :RiyoNaisen, " & vbCrLf & _
                                   "RIYO_FAX11 = :RiyoFax11, " & vbCrLf & _
                                   "RIYO_FAX12 = :RiyoFax12, " & vbCrLf & _
                                   "RIYO_FAX13 = :RiyoFax13, " & vbCrLf & _
                                   "RIYO_YUBIN1 = :RiyoYubin1, " & vbCrLf & _
                                   "RIYO_YUBIN2 = :RiyoYubin2, " & vbCrLf & _
                                   "RIYO_TODO = :RiyoTodo, " & vbCrLf & _
                                   "RIYO_SHIKU = :RiyoShiku, " & vbCrLf & _
                                   "RIYO_BAN = :RiyoBan, " & vbCrLf & _
                                   "RIYO_BUILD = :RiyoBuild, " & vbCrLf & _
                                   "RIYO_LVL = :RiyoLvl, " & vbCrLf & _
                                   "AITE_CD = :AiteCd, " & vbCrLf & _
                                   "ONKYO_NM = :OnkyoNm, " & vbCrLf & _
                                   "ONKYO_TANTO_NM = :OnkyoTantoNm, " & vbCrLf & _
                                   "ONKYO_TEL11 = :OnkyoTel11, " & vbCrLf & _
                                   "ONKYO_TEL12 = :OnkyoTel12, " & vbCrLf & _
                                   "ONKYO_TEL13 = :OnkyoTel13, " & vbCrLf & _
                                   "ONKYO_NAISEN = :OnkyoNaisen, " & vbCrLf & _
                                   "ONKYO_FAX11 = :OnkyoFax11, " & vbCrLf & _
                                   "ONKYO_FAX12 = :OnkyoFax12, " & vbCrLf & _
                                   "ONKYO_FAX13 = :OnkyoFax13, " & vbCrLf & _
                                   "ONKYO_MAIL = :OnkyoMail, " & vbCrLf & _
                                   "TOTAL_RIYO_KIN = :TotalRiyoKin, " & vbCrLf & _
                                   "BIKO = :Biko, " & vbCrLf & _
                                   "RIYO_COM = :RiyoCom, " & vbCrLf & _
                                   "TICKET_ENTER_KBN = :TicketEnterKbn, " & vbCrLf & _
                                   "TICKET_DRINK_KBN = :TicketDrinkKbn, " & vbCrLf & _
                                   "HP_KEISAI = :HpKeisai, " & vbCrLf & _
                                   "JOHO_KOKAI_DT = :JohoKokaiDt, " & vbCrLf & _
                                   "JOHO_KOKAI_TIME = :JohoKokaiTime, " & vbCrLf & _
                                   "KOKAI_DT = :KokaiDt, " & vbCrLf & _
                                   "KOKAI_TIME = :KokaiTime, " & vbCrLf & _
                                   "FINPUT_STS = :FinputSts, " & vbCrLf & _
                                   "UP_DT = current_timestamp, " & vbCrLf & _
                                   "UP_USER_CD = :UpUserCd " & vbCrLf & _
                               "WHERE " & vbCrLf & _
                                   "YOYAKU_NO = :YoyakuNo "
    '2016.11.4 m.hayabuchi MOD Start 課題No.59

    'SQL(請求取消)
    ' 2016.03.03 UPD START↓ h.hagiwara
    'Private strEX05U001_S As String = _
    '                       "update BILLPAY_TBL " & vbCrLf & _
    '                           "SET " & vbCrLf & _
    '                               "SEIKYU_INPUT_FLG = '0', " & vbCrLf & _
    '                               "SEIKYU_IRAI_FLG = '0', " & vbCrLf & _
    '                               "UP_DT = current_timestamp, " & vbCrLf & _
    '                               "UP_USER_CD = :UpUserCd " & vbCrLf & _
    '                           "WHERE " & vbCrLf & _
    '                               "YOYAKU_NO = :YoyakuNo " & _
    '                               "AND SEQ = :Seq "
    Private strEX05U001_S As String = _
                           "update BILLPAY_TBL " & vbCrLf & _
                               "SET " & vbCrLf & _
                                   "SEIKYU_INPUT_FLG = '0', " & vbCrLf & _
                                   "SEIKYU_IRAI_FLG = '0', " & vbCrLf & _
                                   "SEIKYU_DT = NULL, " & vbCrLf & _
                                   "NYUKIN_YOTEI_DT = NULL , " & vbCrLf & _
                                   "NYUKIN_KBN = NULL , " & vbCrLf & _
                                   "UP_DT = current_timestamp, " & vbCrLf & _
                                   "UP_USER_CD = :UpUserCd " & vbCrLf & _
                               "WHERE " & vbCrLf & _
                                   "YOYAKU_NO = :YoyakuNo " & _
                                   "AND SEQ = :Seq "
    ' 2016.03.03 UPD END↑ h.hagiwara

    'SQL(入金取消)
    ' 2016.03.03 UPD START↓ h.hagiwara
    'Private strEX05U001_N As String = _
    '                       "update BILLPAY_TBL " & vbCrLf & _
    '                           "SET " & vbCrLf & _
    '                               "NYUKIN_INPUT_FLG = '0', " & vbCrLf & _
    '                               "UP_DT = current_timestamp, " & vbCrLf & _
    '                               "UP_USER_CD = :UpUserCd " & vbCrLf & _
    '                           "WHERE " & vbCrLf & _
    '                               "YOYAKU_NO = :YoyakuNo " & _
    '                               "AND SEQ = :Seq "
    Private strEX05U001_N As String = _
                           "update BILLPAY_TBL " & vbCrLf & _
                               "SET " & vbCrLf & _
                                   "NYUKIN_INPUT_FLG = '0', " & vbCrLf & _
                                   "NYUKIN_DT = NULL, " & vbCrLf & _
                                   "NYUKIN_KIN = NULL, " & vbCrLf & _
                                   "NYUKIN_LINK_NO = NULL , " & vbCrLf & _
                                   "UP_DT = current_timestamp, " & vbCrLf & _
                                   "UP_USER_CD = :UpUserCd " & vbCrLf & _
                               "WHERE " & vbCrLf & _
                                   "YOYAKU_NO = :YoyakuNo " & _
                                   "AND SEQ = :Seq "
    ' 2016.03.03 UPD END↑ h.hagiwara

    '印刷用
    Private strSelectCertificateSql As String = "SELECT " & vbCrLf &
                                                "t1.yoyaku_no, " & vbCrLf &
                                                "t1.riyo_nm || '　' || t1.sekinin_nm riyo_nm, " & vbCrLf &
                                                "to_char(now(),'yyyy年MM月dd日') 予約受付日, " & vbCrLf &
                                                "to_char(now()+ CAST('+14 days' AS INTERVAL),'yyyy年MM月dd日') 予約有効期限, " & vbCrLf &
                                                "'(' ||(ARRAY['日','月','火','水','木','金','土'])[EXTRACT('dow' FROM now()) + 1] || ')' 曜日 , " & vbCrLf &
                                                "t1.shutsuen_nm, " & vbCrLf &
                                                "CASE " & vbCrLf &
                                                "WHEN t2.利用開始日 = t2.利用終了日 THEN t2.利用開始日 " & vbCrLf &
                                                "ELSE t2.利用開始日 || '～' || t2.利用終了日 " & vbCrLf &
                                                "END 利用日時, " & vbCrLf &
                                                "t2.利用日数, " & vbCrLf &
                                                "CASE t2a.studio_kbn " & vbCrLf &
                                                "WHEN '1' THEN '201st' " & vbCrLf &
                                                "WHEN '2' THEN '202st' " & vbCrLf &
                                                "WHEN '3' THEN 'house lockt' " & vbCrLf &
                                                "END studio_kbn, " & vbCrLf &
                                                "CASE t2a.riyo_keitai " & vbCrLf &
                                                "WHEN '1' THEN '時間貸し' " & vbCrLf &
                                                "WHEN '2' THEN 'lock out' " & vbCrLf &
                                                "END riyo_keitai, " & vbCrLf &
                                                "t3.start_time, " & vbCrLf &
                                                "t3.end_time " & vbCrLf &
                                                "FROM " & vbCrLf &
                                                "yoyaku_tbl t1 " & vbCrLf &
                                                "LEFT JOIN " & vbCrLf &
                                                "ydt_tbl t2a " & vbCrLf &
                                                "ON  t1.yoyaku_no = t2a.yoyaku_no " & vbCrLf &
                                                "LEFT JOIN " & vbCrLf &
                                                "( " & vbCrLf &
                                                "SELECT " & vbCrLf &
                                                "yoyaku_no, " & vbCrLf &
                                                "MIN(yoyaku_dt) 利用開始日, " & vbCrLf &
                                                "MAX(yoyaku_dt) 利用終了日, " & vbCrLf &
                                                "COUNT(yoyaku_dt) 利用日数 " & vbCrLf &
                                                "FROM " & vbCrLf &
                                                "ydt_tbl " & vbCrLf &
                                                "WHERE " & vbCrLf &
                                                "shisetu_kbn = '2' " & vbCrLf &
                                                "AND yoyaku_no = :yoyakuNo " & vbCrLf &
                                                "GROUP BY " & vbCrLf &
                                                "yoyaku_no " & vbCrLf &
                                                ") t2 " & vbCrLf &
                                                "ON  t1.yoyaku_no = t2.yoyaku_no " & vbCrLf &
                                                "LEFT JOIN " & vbCrLf &
                                                "( " & vbCrLf &
                                                "SELECT " & vbCrLf &
                                                "yoyaku_no, " & vbCrLf &
                                                "start_time, " & vbCrLf &
                                                "end_time " & vbCrLf &
                                                "FROM " & vbCrLf &
                                                "ydt_tbl " & vbCrLf &
                                                "WHERE " & vbCrLf &
                                                "yoyaku_no = :yoyakuNo " & vbCrLf &
                                                "ORDER BY " & vbCrLf &
                                                "yoyaku_dt LIMIT 1 " & vbCrLf &
                                                ") t3 " & vbCrLf &
                                                "ON  t1.yoyaku_no = t3.yoyaku_no" & vbCrLf &
                                                "WHERE t1.yoyaku_no = :yoyakuNo"

    '付帯設備利用明細（合算）マスタ取得
    ' 2016.02.02 UPD h.hagiwara 期間判定時に比較日を含むように修正( [ '<' ⇒ '<=' ] , [ '>' ⇒ '>=' ] )
    ' 2015.11.27 UPD START↓ h.hagiwara 利用料取得変更
    ''Private strSelectDetailsAllSql As String = "SELECT" & vbCrLf &
    ''                                                "t1.yoyaku_no" & vbCrLf &
    ''                                                ", t5.riyo_nm" & vbCrLf &
    ''                                                ", CASE WHEN t2.min_ydt = t2.max_ydt THEN TO_CHAR(TO_DATE(t2.min_ydt, 'YYYY/MM/DD'), 'yyyy年MM月dd日')" & vbCrLf &
    ''                                                "  ELSE TO_CHAR(TO_DATE(t2.min_ydt, 'YYYY/MM/DD'), 'yyyy年MM月dd日') || " & vbCrLf &
    ''                                                "'～' || TO_CHAR(TO_DATE(t2.max_ydt, 'YYYY/MM/DD'), 'yyyy年MM月dd日') " & vbCrLf &
    ''                                                "  END usedays " & vbCrLf &
    ''                                                ", t5.saiji_nm" & vbCrLf &
    ''                                                ", t2.daycount" & vbCrLf &
    ''                                                ", t2.sumriyo" & vbCrLf &
    ''                                                ", t6.chosei_kin" & vbCrLf &
    ''                                                ", t3.futai_nm" & vbCrLf &
    ''                                                ", t3.tani" & vbCrLf &
    ''                                                ", t3.tanka" & vbCrLf &
    ''                                                ", SUM(t1.futai_su)" & vbCrLf &
    ''                                                ", SUM(t1.futai_shokei)" & vbCrLf &
    ''                                                ", SUM(t1.futai_chosei)" & vbCrLf &
    ''                                                ", t4.sort" & vbCrLf &
    ''                                                ", t3.sort " & vbCrLf &
    ''                                                "FROM" & vbCrLf &
    ''                                                "friyo_meisai_tbl t1 " & vbCrLf &
    ''                                                "LEFT OUTER JOIN ( " & vbCrLf &
    ''                                                "SELECT" & vbCrLf &
    ''                                                "yoyaku_no" & vbCrLf &
    ''                                                ", MIN(yoyaku_dt) as min_ydt" & vbCrLf &
    ''                                                ", MAX(yoyaku_dt) as max_ydt" & vbCrLf &
    ''                                                ", COUNT(yoyaku_dt) as daycount " & vbCrLf &
    ''                                                ", SUM(riyo_kin) as sumriyo " & vbCrLf &
    ''                                                "FROM" & vbCrLf &
    ''                                                "ydt_tbl " & vbCrLf &
    ''                                                "WHERE" & vbCrLf &
    ''                                                "yoyaku_no = :ReserveNo " & vbCrLf &
    ''                                                "GROUP BY" & vbCrLf &
    ''                                                "yoyaku_no" & vbCrLf &
    ''                                                ") t2 " & vbCrLf &
    ''                                                "ON t1.yoyaku_no = t2.yoyaku_no " & vbCrLf &
    ''                                                "LEFT OUTER JOIN futai_mst t3 " & vbCrLf &
    ''                                                "ON t1.futai_bunrui_cd = t3.bunrui_cd " & vbCrLf &
    ''                                                "AND t1.futai_cd = t3.futai_cd " & vbCrLf &
    ''                                                "AND to_timestamp(t1.yoyaku_dt, 'YYYY/MM/DD') > to_timestamp(t3.kikan_from, 'YYYY/MM/DD')" & vbCrLf &
    ''                                                "AND to_timestamp(t1.yoyaku_dt, 'YYYY/MM/DD') < to_timestamp(t3.kikan_to, 'YYYY/MM/DD')" & vbCrLf &
    ''                                                "AND t3.shisetu_kbn = '2' " & vbCrLf &
    ''                                                "LEFT OUTER JOIN fbunrui_mst t4 " & vbCrLf &
    ''                                                "ON t1.futai_bunrui_cd = t4.bunrui_cd " & vbCrLf &
    ''                                                "AND to_timestamp(t1.yoyaku_dt, 'YYYY/MM/DD') > to_timestamp(t4.kikan_from, 'YYYY/MM/DD')" & vbCrLf &
    ''                                                "AND to_timestamp(t1.yoyaku_dt, 'YYYY/MM/DD') < to_timestamp(t4.kikan_to, 'YYYY/MM/DD')" & vbCrLf &
    ''                                                "AND t4.shisetu_kbn = '2' " & vbCrLf &
    ''                                                "LEFT OUTER JOIN yoyaku_tbl t5 " & vbCrLf &
    ''                                                "ON t1.yoyaku_no = t5.yoyaku_no" & vbCrLf &
    ''                                                "LEFT OUTER JOIN ( " & vbCrLf &
    ''                                                "SELECT" & vbCrLf &
    ''                                                "yoyaku_no" & vbCrLf &
    ''                                                ", SUM(chosei_kin) AS chosei_kin " & vbCrLf &
    ''                                                "FROM" & vbCrLf &
    ''                                                "billpay_tbl " & vbCrLf &
    ''                                                "WHERE" & vbCrLf &
    ''                                                "yoyaku_no = :ReserveNo " & vbCrLf &
    ''                                                "AND seikyu_naiyo IN ('1', '3') " & vbCrLf &
    ''                                                "GROUP BY" & vbCrLf &
    ''                                                "yoyaku_no" & vbCrLf &
    ''                                                ") t6 " & vbCrLf &
    ''                                                "ON t1.yoyaku_no = t6.yoyaku_no " & vbCrLf &
    ''                                                "WHERE" & vbCrLf &
    ''                                                "t1.yoyaku_no = :ReserveNo " & vbCrLf &
    ''                                                "AND t4.notax_flg = :NoTaxFlg " & vbCrLf &
    ''                                                "AND t3.sts = '0' " & vbCrLf &
    ''                                                "AND t4.sts = '0' " & vbCrLf &
    ''                                                "GROUP BY" & vbCrLf &
    ''                                                "t1.yoyaku_no" & vbCrLf &
    ''                                                ", t5.riyo_nm" & vbCrLf &
    ''                                                ", usedays" & vbCrLf &
    ''                                                ", t5.saiji_nm" & vbCrLf &
    ''                                                ", t2.daycount" & vbCrLf &
    ''                                                ", t2.sumriyo" & vbCrLf &
    ''                                                ", t6.chosei_kin" & vbCrLf &
    ''                                                ", t3.futai_nm" & vbCrLf &
    ''                                                ", t3.tani" & vbCrLf &
    ''                                                ", t3.tanka" & vbCrLf &
    ''                                                ", t1.futai_cd" & vbCrLf &
    ''                                                ", t4.sort" & vbCrLf &
    ''                                                ", t3.sort " & vbCrLf &
    ''                                                "ORDER BY" & vbCrLf &
    ''                                                "t4.sort" & vbCrLf &
    ''                                                ", t3.sort"
    ' 2016.07.21 UPD h.hagiwara 付帯設備が存在しない場合の利用料未出力の不具合対応
    'Private strSelectDetailsAllSql As String = "SELECT" & vbCrLf &
    '                                                  " t1.yoyaku_no" & vbCrLf &
    '                                                  ", t5.riyo_nm" & vbCrLf &
    '                                                  ", CASE WHEN t2.min_ydt = t2.max_ydt THEN TO_CHAR(TO_DATE(t2.min_ydt, 'YYYY/MM/DD'), 'yyyy年MM月dd日')" & vbCrLf &
    '                                                  "  ELSE TO_CHAR(TO_DATE(t2.min_ydt, 'YYYY/MM/DD'), 'yyyy年MM月dd日') || " & vbCrLf &
    '                                                  "'～' || TO_CHAR(TO_DATE(t2.max_ydt, 'YYYY/MM/DD'), 'yyyy年MM月dd日') " & vbCrLf &
    '                                                  "  END usedays " & vbCrLf &
    '                                                  ", t5.SHUTSUEN_NM" & vbCrLf &
    '                                                  ", COALESCE(t7.daycount,t2.daycount,0)" & vbCrLf &
    '                                                  ", COALESCE(t7.sumriyo,t2.sumriyo,0)" & vbCrLf &
    '                                                  ", COALESCE(t7.chosei_kin,t6.chosei_kin,0)" & vbCrLf &
    '                                                  ", t3.futai_nm" & vbCrLf &
    '                                                  ", t3.tani" & vbCrLf &
    '                                                  ", t3.tanka" & vbCrLf &
    '                                                  ", SUM(t1.futai_su)" & vbCrLf &
    '                                                  ", SUM(t1.futai_shokei)" & vbCrLf &
    '                                                  ", SUM(t1.futai_chosei)" & vbCrLf &
    '                                                  ", t4.sort" & vbCrLf &
    '                                                  ", t3.sort " & vbCrLf &
    '                                           "FROM" & vbCrLf &
    '                                                  " friyo_meisai_tbl t1 " & vbCrLf &
    '                                             "LEFT OUTER JOIN ( " & vbCrLf &
    '                                                               "SELECT" & vbCrLf &
    '                                                                      "yoyaku_no" & vbCrLf &
    '                                                                      ", MIN(yoyaku_dt) as min_ydt" & vbCrLf &
    '                                                                      ", MAX(yoyaku_dt) as max_ydt" & vbCrLf &
    '                                                                      ", COUNT(yoyaku_dt) as daycount " & vbCrLf &
    '                                                                      ", SUM(riyo_kin) as sumriyo " & vbCrLf &
    '                                                               "FROM" & vbCrLf &
    '                                                                     "ydt_tbl " & vbCrLf &
    '                                                               "WHERE" & vbCrLf &
    '                                                                     "yoyaku_no = :ReserveNo " & vbCrLf &
    '                                                               "GROUP BY" & vbCrLf &
    '                                                                     "yoyaku_no" & vbCrLf &
    '                                                             ") t2 " & vbCrLf &
    '                                                    "ON t1.yoyaku_no = t2.yoyaku_no " & vbCrLf &
    '                                             "LEFT OUTER JOIN futai_mst t3 " & vbCrLf &
    '                                                    "ON t1.futai_bunrui_cd = t3.bunrui_cd " & vbCrLf &
    '                                                   "AND t1.futai_cd = t3.futai_cd " & vbCrLf &
    '                                                   "AND to_timestamp(t1.yoyaku_dt, 'YYYY/MM/DD') >= to_timestamp(t3.kikan_from, 'YYYY/MM/DD')" & vbCrLf &
    '                                                   "AND to_timestamp(t1.yoyaku_dt, 'YYYY/MM/DD') <= to_timestamp(t3.kikan_to, 'YYYY/MM/DD')" & vbCrLf &
    '                                                   "AND t3.shisetu_kbn = '2' " & vbCrLf &
    '                                             "LEFT OUTER JOIN fbunrui_mst t4 " & vbCrLf &
    '                                                    "ON t1.futai_bunrui_cd = t4.bunrui_cd " & vbCrLf &
    '                                                   "AND to_timestamp(t1.yoyaku_dt, 'YYYY/MM/DD') >= to_timestamp(t4.kikan_from, 'YYYY/MM/DD')" & vbCrLf &
    '                                                   "AND to_timestamp(t1.yoyaku_dt, 'YYYY/MM/DD') <= to_timestamp(t4.kikan_to, 'YYYY/MM/DD')" & vbCrLf &
    '                                                   "AND t4.shisetu_kbn = '2' " & vbCrLf &
    '                                             "LEFT OUTER JOIN yoyaku_tbl t5 " & vbCrLf &
    '                                                    "ON t1.yoyaku_no = t5.yoyaku_no" & vbCrLf &
    '                                             "LEFT OUTER JOIN ( " & vbCrLf &
    '                                                               "SELECT" & vbCrLf &
    '                                                                     "yoyaku_no" & vbCrLf &
    '                                                                     ", SUM(chosei_kin) AS chosei_kin " & vbCrLf &
    '                                                               "FROM" & vbCrLf &
    '                                                                     "billpay_tbl " & vbCrLf &
    '                                                               "WHERE" & vbCrLf &
    '                                                                    "yoyaku_no = :ReserveNo " & vbCrLf &
    '                                                                "AND seikyu_naiyo IN ('1', '3') " & vbCrLf &
    '                                                               "GROUP BY" & vbCrLf &
    '                                                                    "yoyaku_no" & vbCrLf &
    '                                                             ") t6 " & vbCrLf &
    '                                                    "ON t1.yoyaku_no = t6.yoyaku_no " & vbCrLf &
    '                                            "  LEFT JOIN(" & vbCrLf &
    '                                                "         SELECT" & vbCrLf &
    '                                                "              t70.yoyaku_no," & vbCrLf &
    '                                                "              SUM(t70.futai_kin) sumriyo," & vbCrLf &
    '                                                "              SUM(t70.futai_chosei) chosei_kin" & vbCrLf &
    '                                                "            , COUNT(t70.yoyaku_dt) as daycount " & vbCrLf &
    '                                                "         FROM  friyo_meisai_tbl t70" & vbCrLf &
    '                                                "         WHERE" & vbCrLf &
    '                                                "               t70.FUTAI_BUNRUI_CD IN ( SELECT  bunrui_cd " & vbCrLf &
    '                                                "                                    FROM    fbunrui_mst t71" & vbCrLf &
    '                                                "                                    WHERE   t71.shisetu_kbn = '2' " & vbCrLf &
    '                                                "                                     AND    to_timestamp(t70.yoyaku_dt, 'YYYY/MM/DD') >= to_timestamp(t71.kikan_from, 'YYYY/MM/DD')" & vbCrLf &
    '                                                "                                     AND    to_timestamp(t70.yoyaku_dt, 'YYYY/MM/DD') <= to_timestamp(t71.kikan_to, 'YYYY/MM/DD')" & vbCrLf &
    '                                                "                                     AND    t71.RIYORYO_FLG = '1' " & vbCrLf &
    '                                                "                                  )" & vbCrLf &
    '                                                "          AND  t70.yoyaku_no = :ReserveNo" & vbCrLf &
    '                                                "         GROUP BY" & vbCrLf &
    '                                                "               t70.yoyaku_no" & vbCrLf &
    '                                            "          ) t7" & vbCrLf &
    '                                                   " ON t1.yoyaku_no = t7.yoyaku_no" & vbCrLf &
    '                                           "WHERE" & vbCrLf &
    '                                                 "t1.yoyaku_no = :ReserveNo " & vbCrLf &
    '                                            "AND  t4.notax_flg = :NoTaxFlg " & vbCrLf &
    '                                            "AND t3.sts = '0' " & vbCrLf &
    '                                            "AND t4.sts = '0' " & vbCrLf &
    '                                            "AND t4.RIYORYO_FLG = '0'" & vbCrLf &
    '                                           "GROUP BY" & vbCrLf &
    '                                                "t1.yoyaku_no" & vbCrLf &
    '                                                ", t5.riyo_nm" & vbCrLf &
    '                                                ", usedays" & vbCrLf &
    '                                                ", t5.SHUTSUEN_NM" & vbCrLf &
    '                                                ", t7.daycount" & vbCrLf &
    '                                                ", t2.daycount" & vbCrLf &
    '                                                ", t7.sumriyo" & vbCrLf &
    '                                                ", t2.sumriyo" & vbCrLf &
    '                                                ", t7.chosei_kin" & vbCrLf &
    '                                                ", t6.chosei_kin" & vbCrLf &
    '                                                ", t3.futai_nm" & vbCrLf &
    '                                                ", t3.tani" & vbCrLf &
    '                                                ", t3.tanka" & vbCrLf &
    '                                                ", t1.futai_cd" & vbCrLf &
    '                                                ", t4.sort" & vbCrLf &
    '                                                ", t3.sort " & vbCrLf &
    '                                           "ORDER BY" & vbCrLf &
    '                                                "t4.sort" & vbCrLf &
    '                                                ", t3.sort"
    ' 2015.11.27 UPD END↑ h.hagiwara 利用料取得変更
    ' --- 2019/08/06 軽減税率対応 Start E.Okuda@Compass ---
    'Private strSelectDetailsAllSql As String = "SELECT" & vbCrLf &
    '                                           "         t2.yoyaku_no" & vbCrLf &
    '                                           "       , t5.riyo_nm" & vbCrLf &
    '                                           "       , CASE WHEN t2.min_ydt = t2.max_ydt THEN TO_CHAR(TO_DATE(t2.min_ydt, 'YYYY/MM/DD'), 'yyyy年MM月dd日')" & vbCrLf &
    '                                           "              ELSE TO_CHAR(TO_DATE(t2.min_ydt, 'YYYY/MM/DD'), 'yyyy年MM月dd日') || " & vbCrLf &
    '                                           "                   '～' || TO_CHAR(TO_DATE(t2.max_ydt, 'YYYY/MM/DD'), 'yyyy年MM月dd日') " & vbCrLf &
    '                                           "         END usedays " & vbCrLf &
    '                                           "      , t5.SHUTSUEN_NM" & vbCrLf &
    '                                           "      , COALESCE(t7.daycount,t2.daycount,0)" & vbCrLf &
    '                                           "      , COALESCE(t7.sumriyo,t2.sumriyo,0)" & vbCrLf &
    '                                           "      , COALESCE(t7.chosei_kin,t6.chosei_kin,0)" & vbCrLf &
    '                                           "      , t3.futai_nm" & vbCrLf &
    '                                           "      , t3.tani" & vbCrLf &
    '                                           "      , t3.tanka" & vbCrLf &
    '                                           "      , SUM(COALESCE(t1.futai_su,0))" & vbCrLf &
    '                                           "      , SUM(COALESCE(t1.futai_shokei,0))" & vbCrLf &
    '                                           "      , SUM(COALESCE(t1.futai_chosei,0))" & vbCrLf &
    '                                           "      , t4.sort" & vbCrLf &
    '                                           "      , t3.sort " & vbCrLf &
    '                                           "FROM" & vbCrLf &
    '                                           "      ( SELECT" & vbCrLf &
    '                                           "             yoyaku_no" & vbCrLf &
    '                                           "           , MIN(yoyaku_dt) as min_ydt" & vbCrLf &
    '                                           "           , MAX(yoyaku_dt) as max_ydt" & vbCrLf &
    '                                           "           , COUNT(yoyaku_dt) as daycount " & vbCrLf &
    '                                           "           , SUM(riyo_kin) as sumriyo " & vbCrLf &
    '                                           "        FROM" & vbCrLf &
    '                                           "             ydt_tbl " & vbCrLf &
    '                                           "        WHERE" & vbCrLf &
    '                                           "             yoyaku_no = :ReserveNo " & vbCrLf &
    '                                           "        GROUP BY" & vbCrLf &
    '                                           "             yoyaku_no" & vbCrLf &
    '                                           "      ) t2 " & vbCrLf &
    '                                           "  LEFT OUTER JOIN friyo_meisai_tbl t1 " & vbCrLf &
    '                                           "         ON t2.yoyaku_no = t1.yoyaku_no " & vbCrLf &
    '                                           "  LEFT OUTER JOIN futai_mst t3 " & vbCrLf &
    '                                           "         ON t1.futai_bunrui_cd = t3.bunrui_cd " & vbCrLf &
    '                                           "        AND t1.futai_cd = t3.futai_cd " & vbCrLf &
    '                                           "        AND to_timestamp(t1.yoyaku_dt, 'YYYY/MM/DD') >= to_timestamp(t3.kikan_from, 'YYYY/MM/DD')" & vbCrLf &
    '                                           "        AND to_timestamp(t1.yoyaku_dt, 'YYYY/MM/DD') <= to_timestamp(t3.kikan_to, 'YYYY/MM/DD')" & vbCrLf &
    '                                           "        AND t3.shisetu_kbn = '2' " & vbCrLf &
    '                                           "        AND t3.sts = '0' " & vbCrLf &
    '                                           "  LEFT OUTER JOIN fbunrui_mst t4 " & vbCrLf &
    '                                           "         ON t1.futai_bunrui_cd = t4.bunrui_cd " & vbCrLf &
    '                                           "        AND to_timestamp(t1.yoyaku_dt, 'YYYY/MM/DD') >= to_timestamp(t4.kikan_from, 'YYYY/MM/DD')" & vbCrLf &
    '                                           "        AND to_timestamp(t1.yoyaku_dt, 'YYYY/MM/DD') <= to_timestamp(t4.kikan_to, 'YYYY/MM/DD')" & vbCrLf &
    '                                           "        AND t4.shisetu_kbn = '2' " & vbCrLf &
    '                                           "        AND t4.sts = '0' " & vbCrLf &
    '                                           "        AND t4.RIYORYO_FLG = '0'" & vbCrLf &
    '                                           "  LEFT OUTER JOIN yoyaku_tbl t5 " & vbCrLf &
    '                                           "         ON t2.yoyaku_no = t5.yoyaku_no" & vbCrLf &
    '                                           "  LEFT OUTER JOIN ( " & vbCrLf &
    '                                           "                    SELECT" & vbCrLf &
    '                                           "                          yoyaku_no" & vbCrLf &
    '                                           "                          , SUM(chosei_kin) AS chosei_kin " & vbCrLf &
    '                                           "                    FROM" & vbCrLf &
    '                                           "                          billpay_tbl " & vbCrLf &
    '                                           "                    WHERE" & vbCrLf &
    '                                           "                         yoyaku_no = :ReserveNo " & vbCrLf &
    '                                           "                     AND seikyu_naiyo IN ('1', '3') " & vbCrLf &
    '                                           "                    GROUP BY" & vbCrLf &
    '                                           "                         yoyaku_no" & vbCrLf &
    '                                           "                  ) t6 " & vbCrLf &
    '                                           "         ON t2.yoyaku_no = t6.yoyaku_no " & vbCrLf &
    '                                           "  LEFT JOIN(" & vbCrLf &
    '                                           "              SELECT" & vbCrLf &
    '                                           "                   t70.yoyaku_no," & vbCrLf &
    '                                           "                   SUM(t70.futai_kin) sumriyo," & vbCrLf &
    '                                           "                   SUM(t70.futai_chosei) chosei_kin" & vbCrLf &
    '                                           "                 , COUNT(t70.yoyaku_dt) as daycount " & vbCrLf &
    '                                           "              FROM  friyo_meisai_tbl t70" & vbCrLf &
    '                                           "              WHERE" & vbCrLf &
    '                                           "                    t70.FUTAI_BUNRUI_CD IN ( SELECT  bunrui_cd " & vbCrLf &
    '                                           "                                         FROM    fbunrui_mst t71" & vbCrLf &
    '                                           "                                         WHERE   t71.shisetu_kbn = '2' " & vbCrLf &
    '                                           "                                          AND    to_timestamp(t70.yoyaku_dt, 'YYYY/MM/DD') >= to_timestamp(t71.kikan_from, 'YYYY/MM/DD')" & vbCrLf &
    '                                           "                                          AND    to_timestamp(t70.yoyaku_dt, 'YYYY/MM/DD') <= to_timestamp(t71.kikan_to, 'YYYY/MM/DD')" & vbCrLf &
    '                                           "                                          AND    t71.RIYORYO_FLG = '1' " & vbCrLf &
    '                                           "                                       )" & vbCrLf &
    '                                           "               AND  t70.yoyaku_no = :ReserveNo" & vbCrLf &
    '                                           "              GROUP BY" & vbCrLf &
    '                                           "                    t70.yoyaku_no" & vbCrLf &
    '                                           "           ) t7" & vbCrLf &
    '                                           "         ON t2.yoyaku_no = t7.yoyaku_no" & vbCrLf &
    '                                           "WHERE" & vbCrLf &
    '                                           "      t2.yoyaku_no = :ReserveNo " & vbCrLf &
    '                                           "  AND COALESCE(t4.notax_flg,'0') = :NoTaxFlg " & vbCrLf &
    '                                           "GROUP BY" & vbCrLf &
    '                                           "       t2.yoyaku_no" & vbCrLf &
    '                                           "     , t5.riyo_nm" & vbCrLf &
    '                                           "     , usedays" & vbCrLf &
    '                                           "     , t5.SHUTSUEN_NM" & vbCrLf &
    '                                           "     , t7.daycount" & vbCrLf &
    '                                           "     , t2.daycount" & vbCrLf &
    '                                           "     , t7.sumriyo" & vbCrLf &
    '                                           "     , t2.sumriyo" & vbCrLf &
    '                                           "     , t7.chosei_kin" & vbCrLf &
    '                                           "     , t6.chosei_kin" & vbCrLf &
    '                                           "     , t3.futai_nm" & vbCrLf &
    '                                           "     , t3.tani" & vbCrLf &
    '                                           "     , t3.tanka" & vbCrLf &
    '                                           "     , t1.futai_cd" & vbCrLf &
    '                                           "     , t4.sort" & vbCrLf &
    '                                           "     , t3.sort " & vbCrLf &
    '                                           "ORDER BY" & vbCrLf &
    '                                           "       t4.sort" & vbCrLf &
    '                                           "     , t3.sort"
    Private strSelectDetailsAllSql As String = "SELECT" & vbCrLf &
                                               "         t2.yoyaku_no" & vbCrLf &
                                               "       , t5.riyo_nm" & vbCrLf &
                                               "       , CASE WHEN t2.min_ydt = t2.max_ydt THEN TO_CHAR(TO_DATE(t2.min_ydt, 'YYYY/MM/DD'), 'yyyy年MM月dd日')" & vbCrLf &
                                               "              ELSE TO_CHAR(TO_DATE(t2.min_ydt, 'YYYY/MM/DD'), 'yyyy年MM月dd日') || " & vbCrLf &
                                               "                   '～' || TO_CHAR(TO_DATE(t2.max_ydt, 'YYYY/MM/DD'), 'yyyy年MM月dd日') " & vbCrLf &
                                               "         END usedays " & vbCrLf &
                                               "      , t5.SHUTSUEN_NM" & vbCrLf &
                                               "      , COALESCE(t7.daycount,t2.daycount,0)" & vbCrLf &
                                               "      , COALESCE(t7.sumriyo,t2.sumriyo,0)" & vbCrLf &
                                               "      , COALESCE(t7.chosei_kin,t6.chosei_kin,0)" & vbCrLf &
                                               "      , t3.futai_nm" & vbCrLf &
                                               "      , t3.tani" & vbCrLf &
                                               "      , t3.tanka" & vbCrLf &
                                               "      , SUM(COALESCE(t1.futai_su,0))" & vbCrLf &
                                               "      , SUM(COALESCE(t1.futai_shokei,0))" & vbCrLf &
                                               "      , SUM(COALESCE(t1.futai_chosei,0))" & vbCrLf &
                                               "      , t4.notax_flg" & vbCrLf &
                                               "      , (CASE WHEN t4.zeiritsu IS NULL THEN t8.tax_ritu ELSE t4.zeiritsu END) as zeiritsu" & vbCrLf &
                                               "      , t4.sort" & vbCrLf &
                                               "      , t3.sort " & vbCrLf &
                                               "FROM" & vbCrLf &
                                               "      ( SELECT" & vbCrLf &
                                               "             yoyaku_no" & vbCrLf &
                                               "           , MIN(yoyaku_dt) as min_ydt" & vbCrLf &
                                               "           , MAX(yoyaku_dt) as max_ydt" & vbCrLf &
                                               "           , COUNT(yoyaku_dt) as daycount " & vbCrLf &
                                               "           , SUM(riyo_kin) as sumriyo " & vbCrLf &
                                               "        FROM" & vbCrLf &
                                               "             ydt_tbl " & vbCrLf &
                                               "        WHERE" & vbCrLf &
                                               "             yoyaku_no = :ReserveNo " & vbCrLf &
                                               "        GROUP BY" & vbCrLf &
                                               "             yoyaku_no" & vbCrLf &
                                               "      ) t2 " & vbCrLf &
                                               "  LEFT OUTER JOIN friyo_meisai_tbl t1 " & vbCrLf &
                                               "         ON t2.yoyaku_no = t1.yoyaku_no " & vbCrLf &
                                               "  LEFT OUTER JOIN futai_mst t3 " & vbCrLf &
                                               "         ON t1.futai_bunrui_cd = t3.bunrui_cd " & vbCrLf &
                                               "        AND t1.futai_cd = t3.futai_cd " & vbCrLf &
                                               "        AND to_timestamp(t1.yoyaku_dt, 'YYYY/MM/DD') >= to_timestamp(t3.kikan_from, 'YYYY/MM/DD')" & vbCrLf &
                                               "        AND to_timestamp(t1.yoyaku_dt, 'YYYY/MM/DD') <= to_timestamp(t3.kikan_to, 'YYYY/MM/DD')" & vbCrLf &
                                               "        AND t3.shisetu_kbn = '2' " & vbCrLf &
                                               "        AND t3.sts = '0' " & vbCrLf &
                                               "  LEFT OUTER JOIN fbunrui_mst t4 " & vbCrLf &
                                               "         ON t1.futai_bunrui_cd = t4.bunrui_cd " & vbCrLf &
                                               "        AND to_timestamp(t1.yoyaku_dt, 'YYYY/MM/DD') >= to_timestamp(t4.kikan_from, 'YYYY/MM/DD')" & vbCrLf &
                                               "        AND to_timestamp(t1.yoyaku_dt, 'YYYY/MM/DD') <= to_timestamp(t4.kikan_to, 'YYYY/MM/DD')" & vbCrLf &
                                               "        AND t4.shisetu_kbn = '2' " & vbCrLf &
                                               "        AND t4.sts = '0' " & vbCrLf &
                                               "        AND t4.RIYORYO_FLG = '0'" & vbCrLf &
                                               "  LEFT OUTER JOIN yoyaku_tbl t5 " & vbCrLf &
                                               "         ON t2.yoyaku_no = t5.yoyaku_no" & vbCrLf &
                                               "  LEFT OUTER JOIN ( " & vbCrLf &
                                               "                    SELECT" & vbCrLf &
                                               "                          yoyaku_no" & vbCrLf &
                                               "                          , SUM(chosei_kin) AS chosei_kin " & vbCrLf &
                                               "                    FROM" & vbCrLf &
                                               "                          billpay_tbl " & vbCrLf &
                                               "                    WHERE" & vbCrLf &
                                               "                         yoyaku_no = :ReserveNo " & vbCrLf &
                                               "                     AND seikyu_naiyo IN ('1', '3') " & vbCrLf &
                                               "                    GROUP BY" & vbCrLf &
                                               "                         yoyaku_no" & vbCrLf &
                                               "                  ) t6 " & vbCrLf &
                                               "         ON t2.yoyaku_no = t6.yoyaku_no " & vbCrLf &
                                               "  LEFT JOIN(" & vbCrLf &
                                               "              SELECT" & vbCrLf &
                                               "                   t70.yoyaku_no," & vbCrLf &
                                               "                   SUM(t70.futai_kin) sumriyo," & vbCrLf &
                                               "                   SUM(t70.futai_chosei) chosei_kin" & vbCrLf &
                                               "                 , COUNT(t70.yoyaku_dt) as daycount " & vbCrLf &
                                               "              FROM  friyo_meisai_tbl t70" & vbCrLf &
                                               "              WHERE" & vbCrLf &
                                               "                    t70.FUTAI_BUNRUI_CD IN ( SELECT  bunrui_cd " & vbCrLf &
                                               "                                         FROM    fbunrui_mst t71" & vbCrLf &
                                               "                                         WHERE   t71.shisetu_kbn = '2' " & vbCrLf &
                                               "                                          AND    to_timestamp(t70.yoyaku_dt, 'YYYY/MM/DD') >= to_timestamp(t71.kikan_from, 'YYYY/MM/DD')" & vbCrLf &
                                               "                                          AND    to_timestamp(t70.yoyaku_dt, 'YYYY/MM/DD') <= to_timestamp(t71.kikan_to, 'YYYY/MM/DD')" & vbCrLf &
                                               "                                          AND    t71.RIYORYO_FLG = '1' " & vbCrLf &
                                               "                                       )" & vbCrLf &
                                               "               AND  t70.yoyaku_no = :ReserveNo" & vbCrLf &
                                               "              GROUP BY" & vbCrLf &
                                               "                    t70.yoyaku_no" & vbCrLf &
                                               "           ) t7" & vbCrLf &
                                               "         ON t2.yoyaku_no = t7.yoyaku_no" & vbCrLf &
                                               "  LEFT JOIN tax_mst t8" & vbCrLf &
                                               "         ON  to_timestamp(t1.yoyaku_dt, 'YYYY/MM/DD') >= to_timestamp(t8.taxs_dt, 'YYYY/MM/DD')" & vbCrLf &
                                               "        AND  to_timestamp(t1.yoyaku_dt, 'YYYY/MM/DD') <= to_timestamp(t8.taxe_dt, 'YYYY/MM/DD')" & vbCrLf &
                                               "WHERE" & vbCrLf &
                                               "      t2.yoyaku_no = :ReserveNo " & vbCrLf &
                                               "  AND COALESCE(t4.notax_flg,'0') = :NoTaxFlg " & vbCrLf &
                                               "GROUP BY" & vbCrLf &
                                               "       t2.yoyaku_no" & vbCrLf &
                                               "     , t5.riyo_nm" & vbCrLf &
                                               "     , usedays" & vbCrLf &
                                               "     , t5.SHUTSUEN_NM" & vbCrLf &
                                               "     , t7.daycount" & vbCrLf &
                                               "     , t2.daycount" & vbCrLf &
                                               "     , t7.sumriyo" & vbCrLf &
                                               "     , t2.sumriyo" & vbCrLf &
                                               "     , t7.chosei_kin" & vbCrLf &
                                               "     , t6.chosei_kin" & vbCrLf &
                                               "     , t3.futai_nm" & vbCrLf &
                                               "     , t3.tani" & vbCrLf &
                                               "     , t3.tanka" & vbCrLf &
                                               "     , t1.futai_cd" & vbCrLf &
                                               "     , t4.notax_flg" & vbCrLf &
                                               "     , t4.zeiritsu" & vbCrLf &
                                               "     , t8.tax_ritu" & vbCrLf &
                                               "     , t4.sort" & vbCrLf &
                                               "     , t3.sort " & vbCrLf &
                                               "ORDER BY" & vbCrLf &
                                               "       t4.sort" & vbCrLf &
                                               "     , t3.sort"
    ' --- 2019/08/06 軽減税率対応 End E.Okuda@Compass ---

    ' 2016.07.21 UPD h.hagiwara 付帯設備が存在しない場合の利用料未出力の不具合対応

    '付帯設備利用明細（日別）マスタ取得
    ' 2016.02.02 UPD h.hagiwara 期間判定時に比較日を含むように修正( [ '<' ⇒ '<=' ] , [ '>' ⇒ '>=' ] )
    ' 2015.11.27 UPD START↓ h.hagiwara 利用料取得変更
    ''Private strSelectDetailsOneDaySql As String = "SELECT" & vbCrLf &
    ''                                                "t1.yoyaku_no" & vbCrLf &
    ''                                                ", t4.riyo_nm" & vbCrLf &
    ''                                                ", TO_CHAR(TO_DATE(t1.yoyaku_dt, 'YYYY/MM/DD'), 'yyyy年MM月dd日') AS yoyaku_dt " & vbCrLf &
    ''                                                ", t4.saiji_nm" & vbCrLf &
    ''                                                ", t5.riyo_kin" & vbCrLf &
    ''                                                ", t2.futai_nm" & vbCrLf &
    ''                                                ", t2.tani" & vbCrLf &
    ''                                                ", t2.tanka" & vbCrLf &
    ''                                                ", t1.futai_su" & vbCrLf &
    ''                                                ", t1.futai_shokei" & vbCrLf &
    ''                                                ", t1.futai_chosei" & vbCrLf &
    ''                                                ", t1.futai_biko" & vbCrLf &
    ''                                                ", t3.sort" & vbCrLf &
    ''                                                ", t2.sort " & vbCrLf &
    ''                                                "FROM" & vbCrLf &
    ''                                                "friyo_meisai_tbl t1 " & vbCrLf &
    ''                                                "LEFT OUTER JOIN futai_mst t2 " & vbCrLf &
    ''                                                "ON t1.futai_bunrui_cd = t2.bunrui_cd " & vbCrLf &
    ''                                                "AND t1.futai_cd = t2.futai_cd " & vbCrLf &
    ''                                                "AND to_timestamp(t1.yoyaku_dt, 'YYYY/MM/DD') > to_timestamp(t2.kikan_from, 'YYYY/MM/DD')" & vbCrLf &
    ''                                                "AND to_timestamp(t1.yoyaku_dt, 'YYYY/MM/DD') < to_timestamp(t2.kikan_to, 'YYYY/MM/DD')" & vbCrLf &
    ''                                                "AND t2.shisetu_kbn = '2' " & vbCrLf &
    ''                                                "LEFT OUTER JOIN fbunrui_mst t3 " & vbCrLf &
    ''                                                "ON t1.futai_bunrui_cd = t3.bunrui_cd " & vbCrLf &
    ''                                                "AND to_timestamp(t1.yoyaku_dt, 'YYYY/MM/DD') > to_timestamp(t3.kikan_from, 'YYYY/MM/DD')" & vbCrLf &
    ''                                                "AND to_timestamp(t1.yoyaku_dt, 'YYYY/MM/DD') < to_timestamp(t3.kikan_to, 'YYYY/MM/DD')" & vbCrLf &
    ''                                                "AND t3.shisetu_kbn = '2' " & vbCrLf &
    ''                                                "LEFT OUTER JOIN yoyaku_tbl t4 " & vbCrLf &
    ''                                                "ON t1.yoyaku_no = t4.yoyaku_no" & vbCrLf &
    ''                                                "LEFT JOIN ydt_tbl t5 " & vbCrLf &
    ''                                                "ON t1.yoyaku_no = t5.yoyaku_no " & vbCrLf &
    ''                                                "AND t1.yoyaku_dt = t5.yoyaku_dt " & vbCrLf &
    ''                                                "WHERE" & vbCrLf &
    ''                                                "t1.yoyaku_no = :ReserveNo " & vbCrLf &
    ''                                                "AND t3.notax_flg = :NoTaxFlg " & vbCrLf &
    ''                                                "AND t2.sts = '0' " & vbCrLf &
    ''                                                "AND t3.sts = '0' " & vbCrLf &
    ''                                                "AND to_timestamp(t1.yoyaku_dt, 'yyyy/MM/dd') = to_timestamp(:ReserveDay, 'yyyy/MM/dd')" & vbCrLf &
    ''                                                "ORDER BY" & vbCrLf &
    ''                                                "t3.sort" & vbCrLf &
    ''                                                ", t2.sort"
    ' 2016.07.21 UPD h.hagiwara 付帯設備が存在しない場合の利用料未出力の不具合対応
    'Private strSelectDetailsOneDaySql As String = "SELECT" & vbCrLf &
    '                                                   "  t1.yoyaku_no" & vbCrLf &
    '                                                   ", t4.riyo_nm" & vbCrLf &
    '                                                   ", TO_CHAR(TO_DATE(t1.yoyaku_dt, 'YYYY/MM/DD'), 'yyyy年MM月dd日') AS yoyaku_dt " & vbCrLf &
    '                                                   ", t4.SHUTSUEN_NM" & vbCrLf &
    '                                                   ", COALESCE(t6.riyo_kin,t5.riyo_kin,0)" & vbCrLf &
    '                                                   ", t2.futai_nm" & vbCrLf &
    '                                                   ", t2.tani" & vbCrLf &
    '                                                   ", t2.tanka" & vbCrLf &
    '                                                   ", t1.futai_su" & vbCrLf &
    '                                                   ", t1.futai_shokei" & vbCrLf &
    '                                                   ", t1.futai_chosei" & vbCrLf &
    '                                                   ", t1.futai_biko" & vbCrLf &
    '                                                   ", t3.sort" & vbCrLf &
    '                                                   ", t2.sort " & vbCrLf &
    '                                                   ", COALESCE(t6.chosei_kin,0)" & vbCrLf &
    '                                              "FROM" & vbCrLf &
    '                                                   "friyo_meisai_tbl t1 " & vbCrLf &
    '                                                "LEFT OUTER JOIN futai_mst t2 " & vbCrLf &
    '                                                       "ON t1.futai_bunrui_cd = t2.bunrui_cd " & vbCrLf &
    '                                                      "AND t1.futai_cd = t2.futai_cd " & vbCrLf &
    '                                                      "AND to_timestamp(t1.yoyaku_dt, 'YYYY/MM/DD') >= to_timestamp(t2.kikan_from, 'YYYY/MM/DD')" & vbCrLf &
    '                                                      "AND to_timestamp(t1.yoyaku_dt, 'YYYY/MM/DD') <= to_timestamp(t2.kikan_to, 'YYYY/MM/DD')" & vbCrLf &
    '                                                      "AND t2.shisetu_kbn = '2' " & vbCrLf &
    '                                                "LEFT OUTER JOIN fbunrui_mst t3 " & vbCrLf &
    '                                                       "ON t1.futai_bunrui_cd = t3.bunrui_cd " & vbCrLf &
    '                                                      "AND to_timestamp(t1.yoyaku_dt, 'YYYY/MM/DD') >= to_timestamp(t3.kikan_from, 'YYYY/MM/DD')" & vbCrLf &
    '                                                      "AND to_timestamp(t1.yoyaku_dt, 'YYYY/MM/DD') <= to_timestamp(t3.kikan_to, 'YYYY/MM/DD')" & vbCrLf &
    '                                                      "AND t3.shisetu_kbn = '2' " & vbCrLf &
    '                                                "LEFT OUTER JOIN yoyaku_tbl t4 " & vbCrLf &
    '                                                       "ON t1.yoyaku_no = t4.yoyaku_no" & vbCrLf &
    '                                                "LEFT JOIN ydt_tbl t5 " & vbCrLf &
    '                                                      " ON t1.yoyaku_no = t5.yoyaku_no " & vbCrLf &
    '                                                     " AND t1.yoyaku_dt = t5.yoyaku_dt " & vbCrLf &
    '                                                 "LEFT JOIN(" & vbCrLf &
    '                                                 "         SELECT" & vbCrLf &
    '                                                 "              t60.yoyaku_no," & vbCrLf &
    '                                                 "              t60.yoyaku_dt," & vbCrLf &
    '                                                 "              SUM(t60.futai_kin) riyo_kin," & vbCrLf &
    '                                                 "              SUM(t60.futai_chosei) chosei_kin" & vbCrLf &
    '                                                 "         FROM  friyo_meisai_tbl t60" & vbCrLf &
    '                                                 "         WHERE" & vbCrLf &
    '                                                 "               t60.FUTAI_BUNRUI_CD IN ( SELECT  bunrui_cd " & vbCrLf &
    '                                                 "                                    FROM    fbunrui_mst t61" & vbCrLf &
    '                                                 "                                    WHERE   t61.shisetu_kbn = '2' " & vbCrLf &
    '                                                 "                                     AND    to_timestamp(t60.yoyaku_dt, 'YYYY/MM/DD') >= to_timestamp(t61.kikan_from, 'YYYY/MM/DD')" & vbCrLf &
    '                                                 "                                     AND    to_timestamp(t60.yoyaku_dt, 'YYYY/MM/DD') <= to_timestamp(t61.kikan_to, 'YYYY/MM/DD')" & vbCrLf &
    '                                                 "                                     AND    t61.RIYORYO_FLG = '1' " & vbCrLf &
    '                                                 "                                  )" & vbCrLf &
    '                                                 "          AND  t60.yoyaku_no = :ReserveNo" & vbCrLf &
    '                                                 "          AND  t60.YOYAKU_DT = :ReserveDay " & vbCrLf &
    '                                                 "         GROUP BY" & vbCrLf &
    '                                                 "               t60.yoyaku_no,t60.yoyaku_dt" & vbCrLf &
    '                                                 "          ) t6" & vbCrLf &
    '                                                   "    ON t1.yoyaku_no = t6.yoyaku_no" & vbCrLf &
    '                                                   "   AND t1.yoyaku_dt = t6.yoyaku_dt" & vbCrLf &
    '                                              "WHERE" & vbCrLf &
    '                                                    "t1.yoyaku_no = :ReserveNo " & vbCrLf &
    '                                               "AND t3.notax_flg = :NoTaxFlg " & vbCrLf &
    '                                               "AND t2.sts = '0' " & vbCrLf &
    '                                               "AND t3.sts = '0' " & vbCrLf &
    '                                               "AND to_timestamp(t1.yoyaku_dt, 'yyyy/MM/dd') = to_timestamp(:ReserveDay, 'yyyy/MM/dd')" & vbCrLf &
    '                                               "AND t3.RIYORYO_FLG = '0'" & vbCrLf &
    '                                              "ORDER BY" & vbCrLf &
    '                                                    "t3.sort" & vbCrLf &
    '                                                    ", t2.sort"
    '' 2015.11.27 UPD END↑ h.hagiwara 利用料取得変更
    ' --- 2019/07/28 軽減税率対応 Start E.Okuda@Compass ---
    Private strSelectDetailsOneDaySql As String = " SELECT" & vbCrLf &
                                                  "       t4.yoyaku_no" & vbCrLf &
                                                  "     , t4.riyo_nm" & vbCrLf &
                                                  "     , TO_CHAR(TO_DATE(t5.yoyaku_dt, 'YYYY/MM/DD'), 'yyyy年MM月dd日') AS yoyaku_dt " & vbCrLf &
                                                  "     , t4.SHUTSUEN_NM" & vbCrLf &
                                                  "     , COALESCE(t6.riyo_kin,t5.riyo_kin,0) as riyo_kin" & vbCrLf &
                                                  "     , t2.futai_nm" & vbCrLf &
                                                  "     , t2.tani" & vbCrLf &
                                                  "     , t2.tanka" & vbCrLf &
                                                  "     , COALESCE(t1.futai_su,0)" & vbCrLf &
                                                  "     , COALESCE(t1.futai_shokei,0) as futai_kin" & vbCrLf &
                                                  "     , COALESCE(t1.futai_chosei,0) as futai_chosei" & vbCrLf &
                                                  "     , t1.futai_biko" & vbCrLf &
                                                  "     , t3.notax_flg" & vbCrLf &
                                                  "     , (CASE WHEN t3.zeiritsu IS NULL THEN t7.tax_ritu ELSE t3.zeiritsu END) as zeiritsu" & vbCrLf &
                                                  "     , t3.sort" & vbCrLf &
                                                  "     , t2.sort " & vbCrLf &
                                                  "     , COALESCE(t6.chosei_kin,0)" & vbCrLf &
                                                  "FROM " & vbCrLf &
                                                  "      yoyaku_tbl t4 " & vbCrLf &
                                                  "  LEFT OUTER JOIN friyo_meisai_tbl t1" & vbCrLf &
                                                  "         ON t4.yoyaku_no = t1.yoyaku_no" & vbCrLf &
                                                  "        AND to_timestamp(:ReserveDay, 'yyyy/MM/dd') = to_timestamp(t1.yoyaku_dt, 'yyyy/MM/dd')" & vbCrLf &
                                                  "  LEFT JOIN ydt_tbl t5 " & vbCrLf &
                                                  "         ON t4.yoyaku_no = t5.yoyaku_no " & vbCrLf &
                                                  "        AND to_timestamp(:ReserveDay, 'yyyy/MM/dd') = to_timestamp(t5.yoyaku_dt, 'yyyy/MM/dd')" & vbCrLf &
                                                  "  LEFT OUTER JOIN futai_mst t2 " & vbCrLf &
                                                  "         ON t1.futai_bunrui_cd = t2.bunrui_cd " & vbCrLf &
                                                  "        AND t1.futai_cd = t2.futai_cd " & vbCrLf &
                                                  "        AND to_timestamp(t1.yoyaku_dt, 'YYYY/MM/DD') >= to_timestamp(t2.kikan_from, 'YYYY/MM/DD')" & vbCrLf &
                                                  "        AND to_timestamp(t1.yoyaku_dt, 'YYYY/MM/DD') <= to_timestamp(t2.kikan_to, 'YYYY/MM/DD')" & vbCrLf &
                                                  "        AND t2.shisetu_kbn = '2' " & vbCrLf &
                                                  "        AND t2.sts = '0' " & vbCrLf &
                                                  "  LEFT OUTER JOIN fbunrui_mst t3 " & vbCrLf &
                                                  "         ON t1.futai_bunrui_cd = t3.bunrui_cd " & vbCrLf &
                                                  "        AND to_timestamp(t1.yoyaku_dt, 'YYYY/MM/DD') >= to_timestamp(t3.kikan_from, 'YYYY/MM/DD')" & vbCrLf &
                                                  "        AND to_timestamp(t1.yoyaku_dt, 'YYYY/MM/DD') <= to_timestamp(t3.kikan_to, 'YYYY/MM/DD')" & vbCrLf &
                                                  "        AND t3.shisetu_kbn = '2' " & vbCrLf &
                                                  "        AND t3.sts = '0' " & vbCrLf &
                                                  "        AND t3.RIYORYO_FLG = '0'" & vbCrLf &
                                                  "  LEFT JOIN(" & vbCrLf &
                                                  "            SELECT" & vbCrLf &
                                                  "                 t60.yoyaku_no," & vbCrLf &
                                                  "                 t60.yoyaku_dt," & vbCrLf &
                                                  "                 SUM(t60.futai_kin) riyo_kin," & vbCrLf &
                                                  "                 SUM(t60.futai_chosei) chosei_kin" & vbCrLf &
                                                  "            FROM  friyo_meisai_tbl t60" & vbCrLf &
                                                  "            WHERE" & vbCrLf &
                                                  "                  t60.FUTAI_BUNRUI_CD IN ( SELECT  bunrui_cd " & vbCrLf &
                                                  "                                       FROM    fbunrui_mst t61" & vbCrLf &
                                                  "                                       WHERE   t61.shisetu_kbn = '2' " & vbCrLf &
                                                  "                                        AND    to_timestamp(t60.yoyaku_dt, 'YYYY/MM/DD') >= to_timestamp(t61.kikan_from, 'YYYY/MM/DD')" & vbCrLf &
                                                  "                                        AND    to_timestamp(t60.yoyaku_dt, 'YYYY/MM/DD') <= to_timestamp(t61.kikan_to, 'YYYY/MM/DD')" & vbCrLf &
                                                  "                                        AND    t61.RIYORYO_FLG = '1' " & vbCrLf &
                                                  "                                     )" & vbCrLf &
                                                  "             AND  t60.yoyaku_no = :ReserveNo" & vbCrLf &
                                                  "             AND  t60.YOYAKU_DT = :ReserveDay " & vbCrLf &
                                                  "            GROUP BY" & vbCrLf &
                                                  "                  t60.yoyaku_no,t60.yoyaku_dt" & vbCrLf &
                                                  "             ) t6" & vbCrLf &
                                                  "         ON t4.yoyaku_no = t6.yoyaku_no" & vbCrLf &
                                                  "        AND t5.yoyaku_dt = t6.yoyaku_dt" & vbCrLf &
                                                  " LEFT JOIN tax_mst t7" & vbCrLf &
                                                  "       ON  to_timestamp(t1.yoyaku_dt, 'YYYY/MM/DD') >= to_timestamp(t7.taxs_dt, 'YYYY/MM/DD')" & vbCrLf &
                                                  "      AND  to_timestamp(t1.yoyaku_dt, 'YYYY/MM/DD') <= to_timestamp(t7.taxe_dt, 'YYYY/MM/DD')" & vbCrLf &
                                                  "WHERE" & vbCrLf &
                                                  "      t4.yoyaku_no = :ReserveNo " & vbCrLf &
                                                  " AND COALESCE(t3.notax_flg,'0') = :NoTaxFlg " & vbCrLf &
                                                  "ORDER BY" & vbCrLf &
                                                  "        t3.sort" & vbCrLf &
                                                  "      , t2.sort"
    ' 2016.07.21 UPD h.hagiwara 付帯設備が存在しない場合の利用料未出力の不具合対応
    'Private strSelectDetailsOneDaySql As String = " SELECT" & vbCrLf &
    '                                              "       t4.yoyaku_no" & vbCrLf &
    '                                              "     , t4.riyo_nm" & vbCrLf &
    '                                              "     , TO_CHAR(TO_DATE(t5.yoyaku_dt, 'YYYY/MM/DD'), 'yyyy年MM月dd日') AS yoyaku_dt " & vbCrLf &
    '                                              "     , t4.SHUTSUEN_NM" & vbCrLf &
    '                                              "     , COALESCE(t6.riyo_kin,t5.riyo_kin,0)" & vbCrLf &
    '                                              "     , t2.futai_nm" & vbCrLf &
    '                                              "     , t2.tani" & vbCrLf &
    '                                              "     , t2.tanka" & vbCrLf &
    '                                              "     , COALESCE(t1.futai_su,0)" & vbCrLf &
    '                                              "     , COALESCE(t1.futai_shokei,0)" & vbCrLf &
    '                                              "     , COALESCE(t1.futai_chosei,0)" & vbCrLf &
    '                                              "     , t1.futai_biko" & vbCrLf &
    '                                              "     , t3.sort" & vbCrLf &
    '                                              "     , t2.sort " & vbCrLf &
    '                                              "     , COALESCE(t6.chosei_kin,0)" & vbCrLf &
    '                                              "FROM " & vbCrLf &
    '                                              "      yoyaku_tbl t4 " & vbCrLf &
    '                                              "  LEFT OUTER JOIN friyo_meisai_tbl t1" & vbCrLf &
    '                                              "         ON t4.yoyaku_no = t1.yoyaku_no" & vbCrLf &
    '                                              "        AND to_timestamp(:ReserveDay, 'yyyy/MM/dd') = to_timestamp(t1.yoyaku_dt, 'yyyy/MM/dd')" & vbCrLf &
    '                                              "  LEFT JOIN ydt_tbl t5 " & vbCrLf &
    '                                              "         ON t4.yoyaku_no = t5.yoyaku_no " & vbCrLf &
    '                                              "        AND to_timestamp(:ReserveDay, 'yyyy/MM/dd') = to_timestamp(t5.yoyaku_dt, 'yyyy/MM/dd')" & vbCrLf &
    '                                              "  LEFT OUTER JOIN futai_mst t2 " & vbCrLf &
    '                                              "         ON t1.futai_bunrui_cd = t2.bunrui_cd " & vbCrLf &
    '                                              "        AND t1.futai_cd = t2.futai_cd " & vbCrLf &
    '                                              "        AND to_timestamp(t1.yoyaku_dt, 'YYYY/MM/DD') >= to_timestamp(t2.kikan_from, 'YYYY/MM/DD')" & vbCrLf &
    '                                              "        AND to_timestamp(t1.yoyaku_dt, 'YYYY/MM/DD') <= to_timestamp(t2.kikan_to, 'YYYY/MM/DD')" & vbCrLf &
    '                                              "        AND t2.shisetu_kbn = '2' " & vbCrLf &
    '                                              "        AND t2.sts = '0' " & vbCrLf &
    '                                              "  LEFT OUTER JOIN fbunrui_mst t3 " & vbCrLf &
    '                                              "         ON t1.futai_bunrui_cd = t3.bunrui_cd " & vbCrLf &
    '                                              "        AND to_timestamp(t1.yoyaku_dt, 'YYYY/MM/DD') >= to_timestamp(t3.kikan_from, 'YYYY/MM/DD')" & vbCrLf &
    '                                              "        AND to_timestamp(t1.yoyaku_dt, 'YYYY/MM/DD') <= to_timestamp(t3.kikan_to, 'YYYY/MM/DD')" & vbCrLf &
    '                                              "        AND t3.shisetu_kbn = '2' " & vbCrLf &
    '                                              "        AND t3.sts = '0' " & vbCrLf &
    '                                              "        AND t3.RIYORYO_FLG = '0'" & vbCrLf &
    '                                              "  LEFT JOIN(" & vbCrLf &
    '                                              "            SELECT" & vbCrLf &
    '                                              "                 t60.yoyaku_no," & vbCrLf &
    '                                              "                 t60.yoyaku_dt," & vbCrLf &
    '                                              "                 SUM(t60.futai_kin) riyo_kin," & vbCrLf &
    '                                              "                 SUM(t60.futai_chosei) chosei_kin" & vbCrLf &
    '                                              "            FROM  friyo_meisai_tbl t60" & vbCrLf &
    '                                              "            WHERE" & vbCrLf &
    '                                              "                  t60.FUTAI_BUNRUI_CD IN ( SELECT  bunrui_cd " & vbCrLf &
    '                                              "                                       FROM    fbunrui_mst t61" & vbCrLf &
    '                                              "                                       WHERE   t61.shisetu_kbn = '2' " & vbCrLf &
    '                                              "                                        AND    to_timestamp(t60.yoyaku_dt, 'YYYY/MM/DD') >= to_timestamp(t61.kikan_from, 'YYYY/MM/DD')" & vbCrLf &
    '                                              "                                        AND    to_timestamp(t60.yoyaku_dt, 'YYYY/MM/DD') <= to_timestamp(t61.kikan_to, 'YYYY/MM/DD')" & vbCrLf &
    '                                              "                                        AND    t61.RIYORYO_FLG = '1' " & vbCrLf &
    '                                              "                                     )" & vbCrLf &
    '                                              "             AND  t60.yoyaku_no = :ReserveNo" & vbCrLf &
    '                                              "             AND  t60.YOYAKU_DT = :ReserveDay " & vbCrLf &
    '                                              "            GROUP BY" & vbCrLf &
    '                                              "                  t60.yoyaku_no,t60.yoyaku_dt" & vbCrLf &
    '                                              "             ) t6" & vbCrLf &
    '                                              "         ON t4.yoyaku_no = t6.yoyaku_no" & vbCrLf &
    '                                              "        AND t5.yoyaku_dt = t6.yoyaku_dt" & vbCrLf &
    '                                              "WHERE" & vbCrLf &
    '                                              "      t4.yoyaku_no = :ReserveNo " & vbCrLf &
    '                                              " AND COALESCE(t3.notax_flg,'0') = :NoTaxFlg " & vbCrLf &
    '                                              "ORDER BY" & vbCrLf &
    '                                              "        t3.sort" & vbCrLf &
    '                                              "      , t2.sort"
    ' 2016.07.21 UPD h.hagiwara 付帯設備が存在しない場合の利用料未出力の不具合対応
    ' --- 2019/07/28 軽減税率対応 End E.Okuda@Compass ---

    '利用確認セレクト
    Private setSelectStudio As String =
                                    " SELECT" &
                                        " t1.yoyaku_no," &
                                        " COALESCE(t1.riyo_nm,'') || '　' || COALESCE(t1.sekinin_nm,'') 利用者名," &
                                        " to_char(now(),'yyyy年MM月dd日') 予約受付日," &
                                        " to_char(now()+ CAST('+14 days' AS INTERVAL),'yyyy年MM月dd日') 予約有効期限," &
                                        " '(' ||(ARRAY['日','月','火','水','木','金','土'])[EXTRACT('dow' FROM now()) + 1] || ')' 曜日 ," &
                                        " t1.shutsuen_nm," &
                                    " CASE" &
                                        " WHEN t2.利用開始日 = t2.利用終了日 THEN " &
                                                " to_char(CAST(t2.利用開始日 as timestamp),'yyyy年MM月dd日') || '(' ||(ARRAY['日','月','火','水','木','金','土'])[EXTRACT('dow' FROM (SELECT CAST(t2.利用開始日 as timestamp))) + 1] || ')'" &
                                        " ELSE " &
                                                " to_char(CAST(t2.利用開始日 as timestamp),'yyyy年MM月dd日') || '(' ||(ARRAY['日','月','火','水','木','金','土'])[EXTRACT('dow' FROM (SELECT CAST(t2.利用開始日 as timestamp))) + 1] || ')' || " &
                                                " '～' || to_char(CAST(t2.利用終了日 as timestamp),'yyyy年MM月dd日') || '(' ||(ARRAY['日','月','火','水','木','金','土'])[EXTRACT('dow' FROM(SELECT CAST(t2.利用終了日 as timestamp))) + 1] || ')'" &
                                    "  END 利用日時," &
                                        " t2.利用日数," &
                                    " CASE t2a.studio_kbn" &
                                        " WHEN '1' THEN '201st'" &
                                        " WHEN '2' THEN '202st'" &
                                        " WHEN '3' THEN 'house lockt'" &
                                    " END studio_kbn," &
                                    " CASE t2a.riyo_keitai" &
                                        " WHEN '1' THEN '時間貸し'" &
                                        " WHEN '2' THEN 'lock out'" &
                                    " END riyo_keitai," &
                                        " t3.start_time," &
                                        " t3.end_time" &
                                    " FROM" &
                                        " yoyaku_tbl t1" &
                                    " LEFT JOIN" &
                                        " ydt_tbl t2a" &
                                    " ON  t1.yoyaku_no = t2a.yoyaku_no" &
                                        " LEFT JOIN" &
                                            " (" &
                                                " SELECT" &
                                                    " yoyaku_no," &
                                                    " MIN(yoyaku_dt) 利用開始日," &
                                                    " MAX(yoyaku_dt) 利用終了日," &
                                                   "  COUNT(yoyaku_dt) 利用日数" &
                                                " FROM" &
                                                    " ydt_tbl" &
                                                " WHERE" &
                                                   "  shisetu_kbn = '2'" &
                                                " AND yoyaku_no = :yoyakuNo" &
                                                " GROUP BY" &
                                                   "  yoyaku_no" &
                                            " ) t2" &
                                        " ON  t1.yoyaku_no = t2.yoyaku_no" &
                                        " LEFT JOIN" &
                                            " (" &
                                                " SELECT" &
                                                    " yoyaku_no," &
                                                   "  start_time," &
                                                    " end_time" &
                                               "  FROM" &
                                                    " ydt_tbl" &
                                                " WHERE" &
                                                    " yoyaku_no = :yoyakuNo" &
                                                " ORDER BY" &
                                                   "  yoyaku_dt LIMIT 1" &
                                            " ) t3" &
                                        " ON  t1.yoyaku_no = t3.yoyaku_no"
    '祝祭日マスタ存在チェック
    Private strHolidayCntSql As String = "SELECT  COUNT(*) " & vbCrLf &
                                         "FROM HOLIDAY_MST " & vbCrLf &
                                         "WHERE HOLIDAY_DT = :HOLIDAY_DT "
    '税抜フラグ取得用
    Private strNoTaxFlgGetSql As String = "SELECT COALESCE(MAX(notax_flg),'0') " & vbCrLf & _
                                          "FROM FBUNRUI_MST " & vbCrLf & _
                                          "WHERE  SHISETU_KBN = '2' " & vbCrLf & _
                                          " AND   SHUKEI_GRP  = :SHUKEI_GRP " & vbCrLf & _
                                          " AND   :KIKAN_DT BETWEEN KIKAN_FROM AND KIKAN_TO "
    '相手先情報取得用 2014.11.13 ADD h.hagiwara
    Private strAitesakiGetSql As String = "SELECT  POST_BANGO , ADD1 , ADD2 , AITE_NM " & vbCrLf & _
                                          "FROM AITESAKI_MST " & vbCrLf & _
                                          "WHERE  AITE_CD = :AITE_CD "

    ' 2015.12.16 ADD START↓ h.hagiwara 
    ' 料金マスタ情報取得用
    Private strRyokinGetSql As String = "SELECT COUNT(*) " & vbCrLf & _
                                          "FROM RIYORYO_MST " & vbCrLf & _
                                          "WHERE :RIYOUBI BETWEEN KIKAN_FROM AND KIKAN_TO " & vbCrLf & _
                                          " AND  STS = '0' " & vbCrLf & _
                                          " AND  SHISETU_KBN = '2' "
    ' 倍率マスタ情報取得用
    Private strBairituGetSql As String = "SELECT COUNT(*) " & vbCrLf & _
                                          "FROM BAIRITU_MST " & vbCrLf & _
                                          "WHERE :RIYOUBI BETWEEN KIKAN_FROM AND KIKAN_TO " & vbCrLf & _
                                          " AND  STS = '0' " & vbCrLf & _
                                          " AND  SHISETU_KBN = '2' "
    ' 付帯設備マスタ情報取得用
    Private strFutaisetubiGetSql As String = "SELECT COUNT(*) " & vbCrLf & _
                                          "FROM FUTAI_MST " & vbCrLf & _
                                          "WHERE :RIYOUBI BETWEEN KIKAN_FROM AND KIKAN_TO " & vbCrLf & _
                                          " AND  STS = '0' " & vbCrLf & _
                                          " AND  SHISETU_KBN = '2' "
    ' 2015.12.16 ADD END↑ h.hagiwara

    '2016.11.2 MOD m.hayabuchi 課題No.59 携帯番号を一枠(RIYO_TEL21)に統合
    ' 2016.08.04 ADD START e.watanabe 課題No.58
    '責任者メールアドレス・携帯電話番号情報取得
    Private strEX04S005 As String =
                           "SELECT" & vbCrLf & _
                           "    SEKININ_MAIL, RIYO_TEL21" & vbCrLf & _
                           "FROM YOYAKU_TBL" & vbCrLf & _
                           "WHERE COALESCE(KAKUTEI_DT,'Null') = " & vbCrLf & _
                           "(" & vbCrLf & _
                           "    SELECT MAX(COALESCE(KAKUTEI_DT,'Null'))" & vbCrLf & _
                           "    FROM YOYAKU_TBL" & vbCrLf & _
                           "    WHERE SEKININ_NM = :SekininNm" & vbCrLf & _
                           "    AND RIYOSHA_CD = :RiyoshaCd" & vbCrLf & _
                           ")" & vbCrLf & _
                           "AND SEKININ_NM = :SekininNm"
    ' 2016.08.04 ADD END e.watanabe 課題No.58

    ' --- 2019/08/02 軽減税率対応 Start E.Okuda@Compass ---
    Private strSelectTaxRateSql =
                           "SELECT" & vbCrLf &
                           "    TAX_RITU, REDUCED_RATE" & vbCrLf &
                           "FROM TAX_MST" & vbCrLf &
                           "WHERE :RiyouStart BETWEEN TAXS_DT AND TAXE_DT" & vbCrLf &
                           "OR :RiyouEnd BETWEEN TAXS_DT AND TAXE_DT"
    ' --- 2019/08/02 軽減税率対応 End E.Okuda@Compass ---

    ''' <summary>
    ''' 請求情報取得
    ''' </summary>
    ''' <param name="Adapter"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataEXTC0102"></param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>請求情報を取得するSQLの作成
    ''' <para>作成情報：2015/08/25 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function selectBillReq(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByRef Cn As NpgsqlConnection, _
                                       ByRef dataEXTC0102 As DataEXTC0102) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Try
            'SQL文(SELECT)
            Dim Cmd As New NpgsqlCommand
            Cmd.Connection = Cn
            Cmd.CommandText = strEX05S002
            Adapter.SelectCommand = Cmd
            With Adapter.SelectCommand.Parameters
                .Add(New NpgsqlParameter(":YoyakuNo", NpgsqlTypes.NpgsqlDbType.Varchar))
            End With
            With Adapter.SelectCommand
                .Parameters(0).Value = dataEXTC0102.PropStrYoyakuNo
            End With

            Dim ds As New DataSet
            Adapter.Fill(ds, "BILLPAY_TBL")
            dataEXTC0102.PropDsBillReq = ds

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常終了
            Return True
        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
            '例外処理
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 入金情報取得
    ''' </summary>
    ''' <param name="Adapter"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataEXTC0102"></param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>入金情報を取得するSQLの作成
    ''' <para>作成情報：2015/09/02 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function selectNyukin(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByRef Cn As NpgsqlConnection, _
                                       ByRef dataEXTC0102 As DataEXTC0102, _
                                       ByVal StrLink As String, _
                                       ByVal StrIraiNo As String,
                                       ByVal IntLineNo As Integer) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Try
            'SQL文(SELECT)
            Dim Cmd As New NpgsqlCommand
            Cmd.Connection = Cn
            Dim strSQL As New StringBuilder
            strSQL.Append(strEX41S001)
            If String.IsNullOrEmpty(StrLink) = False Then
                ' strSQL.AppendLine(String.Format("AND yoyaku_no = '{0}' ", dataEXTC0102.PropStrYoyakuNo))                            ' 2016.02.12 DEL h.hagiwara
                strSQL.AppendLine(String.Format("AND nyukin_link_no = '{0}' ", StrLink))
            Else
                strSQL.AppendLine(String.Format("AND yoyaku_no = '{0}' ", dataEXTC0102.PropStrYoyakuNo))
                strSQL.AppendLine(String.Format("AND seikyu_irai_no = '{0}' ", StrIraiNo))
                ' 2016.01.14 ADD START↓ y.morooka 入金情報複数件の対応
                strSQL.AppendLine(String.Format("AND nyukin_link_no = (select MIN(nyukin_link_no) FROM exas_nyukin_tbl WHERE nyukin_dt = (select MIN(nyukin_dt) FROM exas_nyukin_tbl WHERE seikyu_irai_no = '{0}') and seikyu_irai_no = '{1}')", StrIraiNo, StrIraiNo))
                ' 2016.01.14 ADD END↑ y.morooka 入金情報複数件の対応
                ' strSQL.AppendLine("AND nyukin_link_no is null ")                    ' 2015.11.13 DEL h.hagiwara
            End If

            Dim dtBillPay As DataTable = dataEXTC0102.PropDsBillReq.Tables("BILLPAY_TBL")
            Cmd.CommandText = strSQL.ToString
            Adapter.SelectCommand = Cmd

            With Adapter.SelectCommand.Parameters
                .Add(New NpgsqlParameter(":LineNo", NpgsqlTypes.NpgsqlDbType.Integer))
            End With
            With Adapter.SelectCommand
                .Parameters(0).Value = IntLineNo
            End With

            Dim table As New DataTable
            Adapter.Fill(table)
            If dataEXTC0102.PropDtExasNyukin Is Nothing Then
                If table.Rows.Count > 0 Then
                    dataEXTC0102.PropDtExasNyukin = table
                Else
                    Dim row As DataRow = table.NewRow()
                    row("line_no") = IntLineNo
                    table.Rows.Add(row)
                    dataEXTC0102.PropDtExasNyukin = table
                End If
            Else
                If table.Rows.Count > 0 Then
                    dataEXTC0102.PropDtExasNyukin.ImportRow(table.Rows(0))
                Else
                    Dim row As DataRow = dataEXTC0102.PropDtExasNyukin.NewRow()
                    row("line_no") = IntLineNo
                    dataEXTC0102.PropDtExasNyukin.Rows.Add(row)
                End If
            End If

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常終了
            Return True
        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
            '例外処理
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 承認依頼取得
    ''' </summary>
    ''' <param name="Adapter"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataEXTC0102"></param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>請求情報を取得するSQLの作成
    ''' <para>作成情報：2015/08/25 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function selectApprovalReq(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByRef Cn As NpgsqlConnection, _
                                       ByRef dataEXTC0102 As DataEXTC0102) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Try
            'SQL文(SELECT)
            Dim Cmd As New NpgsqlCommand
            Cmd.Connection = Cn
            Cmd.CommandText = strEX05S006
            Adapter.SelectCommand = Cmd
            With Adapter.SelectCommand.Parameters
                .Add(New NpgsqlParameter(":YoyakuNo", NpgsqlTypes.NpgsqlDbType.Varchar))
            End With
            With Adapter.SelectCommand
                .Parameters(0).Value = dataEXTC0102.PropStrYoyakuNo
            End With

            Dim ds As New DataSet
            Adapter.Fill(ds, "IRAI_RIREKI_TBL")
            dataEXTC0102.PropDsApproval = ds

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

    ''' <summary>
    ''' 承認記録取得
    ''' </summary>
    ''' <param name="Adapter"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataEXTC0102"></param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>請求情報を取得するSQLの作成
    ''' <para>作成情報：2015/08/25 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function selectApprovalRes(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByRef Cn As NpgsqlConnection, _
                                       ByRef dataEXTC0102 As DataEXTC0102) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Try
            'SQL文(SELECT)
            Dim Cmd As New NpgsqlCommand
            Cmd.Connection = Cn
            Cmd.CommandText = strEX05S007
            Adapter.SelectCommand = Cmd
            With Adapter.SelectCommand.Parameters
                .Add(New NpgsqlParameter(":YoyakuNo", NpgsqlTypes.NpgsqlDbType.Varchar))
            End With
            With Adapter.SelectCommand
                .Parameters(0).Value = dataEXTC0102.PropStrYoyakuNo
            End With

            Dim ds = dataEXTC0102.PropDsApproval
            Adapter.Fill(ds, "CHECK_RIREKI_TBL")
            dataEXTC0102.PropDsApproval = ds

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

    ''' <summary>
    ''' 付帯情報取得
    ''' </summary>
    ''' <param name="Adapter"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataEXTC0102"></param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>付帯情報取得するSQLの作成
    ''' <para>作成情報：2015/08/27 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function selectFutaiHeadder(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByRef Cn As NpgsqlConnection, _
                                       ByRef dataEXTC0102 As DataEXTC0102) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Try
            'SQL文(SELECT)
            Dim Cmd As New NpgsqlCommand
            Dim ht As New Hashtable
            If dataEXTC0102.PropHtFutai IsNot Nothing Then
                ht = dataEXTC0102.PropHtFutai
            End If
            Cmd.Connection = Cn

            For Each dataRiyobi As CommonDataRiyobi In dataEXTC0102.PropListRiyobi
                If ht.ContainsKey(dataRiyobi.PropStrYoyakuDt) = False Then
                    Dim Table As New DataTable()
                    Cmd.CommandText = strEX05S004
                    Adapter.SelectCommand = Cmd
                    With Adapter.SelectCommand.Parameters
                        .Add(New NpgsqlParameter(":YoyakuNo", NpgsqlTypes.NpgsqlDbType.Varchar))
                        .Add(New NpgsqlParameter(":Riyobi", NpgsqlTypes.NpgsqlDbType.Varchar))
                        .Add(New NpgsqlParameter(":ShisetuKbn", NpgsqlTypes.NpgsqlDbType.Varchar))
                        .Add(New NpgsqlParameter(":SYANAI", NpgsqlTypes.NpgsqlDbType.Varchar))
                    End With
                    With Adapter.SelectCommand
                        .Parameters(0).Value = dataEXTC0102.PropStrYoyakuNo
                        .Parameters(1).Value = dataRiyobi.PropStrYoyakuDt
                        .Parameters(2).Value = dataEXTC0102.PropStrShisetuKbn
                        .Parameters(3).Value = dataEXTC0102.PropStrKashiKind
                    End With
                    Adapter.Fill(Table)
                    If Table.Rows.Count > 0 Then
                        ht.Add(dataRiyobi.PropStrYoyakuDt, Table)
                    Else
                        Table = New DataTable()
                        Cmd.CommandText = strEX05S097
                        Adapter.Fill(Table)
                        If Table.Rows.Count = 0 Then
                            '付帯設備初期情報が存在しない場合、仮行を設定する
                            Dim row As DataRow = Table.NewRow
                            row("yoyaku_no") = ""
                            row("yoyaku_dt") = dataRiyobi.PropStrYoyakuDt
                            row("total_fuzoku_kin") = 0
                            row("fuzoku_nm") = ""
                            row("tax_kin") = 0                                      ' 2015.11.13 ADD h.hagiwara
                            Table.Rows.Add(row)
                        End If
                        ht.Add(dataRiyobi.PropStrYoyakuDt, Table)
                    End If
                End If
            Next
            dataEXTC0102.PropHtFutai = ht

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

    ''' <summary>
    ''' 付帯情報取得
    ''' </summary>
    ''' <param name="Adapter"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataEXTC0102"></param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>付帯情報取得するSQLの作成
    ''' <para>作成情報：2015/08/27 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function selectFutaiDetail(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByRef Cn As NpgsqlConnection, _
                                       ByRef dataEXTC0102 As DataEXTC0102) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Try
            'SQL文(SELECT)
            Dim ht As New Hashtable
            Dim Cmd As New NpgsqlCommand
            If dataEXTC0102.PropHtFutaiDetail IsNot Nothing Then
                ht = dataEXTC0102.PropHtFutaiDetail
            End If

            Cmd.Connection = Cn
            Cmd.CommandText = strEX05S098
            Adapter.SelectCommand = Cmd
            With Adapter.SelectCommand.Parameters
                .Add(New NpgsqlParameter(":YoyakuNo", NpgsqlTypes.NpgsqlDbType.Varchar))
                .Add(New NpgsqlParameter(":ShisetuKbn", NpgsqlTypes.NpgsqlDbType.Varchar))
                .Add(New NpgsqlParameter(":Riyobi", NpgsqlTypes.NpgsqlDbType.Varchar))
                .Add(New NpgsqlParameter(":SYANAI", NpgsqlTypes.NpgsqlDbType.Varchar))
            End With
            For Each dataRiyobi As CommonDataRiyobi In dataEXTC0102.PropListRiyobi
                If ht.ContainsKey(dataRiyobi.PropStrYoyakuDt) = False Then
                    With Adapter.SelectCommand
                        .Parameters(0).Value = dataEXTC0102.PropStrYoyakuNo
                        .Parameters(1).Value = dataEXTC0102.PropStrShisetuKbn
                        .Parameters(2).Value = dataRiyobi.PropStrYoyakuDt
                        .Parameters(3).Value = dataEXTC0102.PropStrKashiKind
                    End With
                    Dim Table As New DataTable()
                    Adapter.Fill(Table)
                    If Table.Rows.Count > 0 Then
                        ht.Add(dataRiyobi.PropStrYoyakuDt, Table)
                        Debug.WriteLine("1:" + dataRiyobi.PropStrYoyakuDt + ":" + Table.Rows.Count.ToString)
                    Else
                        ' 2015.12.25 UPD START↓ h.hagiwara
                        'Table = New DataTable()
                        'Cmd.CommandText = strEX05S096
                        'Adapter.Fill(Table)
                        'ht.Add(dataRiyobi.PropStrYoyakuDt, Table)
                        'Debug.WriteLine("2:" + dataRiyobi.PropStrYoyakuDt + ":" + Table.Rows.Count.ToString)
                        If dataEXTC0102.PropStrYoyakuSts = YOYAKU_STS_KARI Then
                            Table = New DataTable()
                            Cmd.CommandText = strEX05S096
                            Adapter.Fill(Table)
                            ht.Add(dataRiyobi.PropStrYoyakuDt, Table)
                            Debug.WriteLine("2:" + dataRiyobi.PropStrYoyakuDt + ":" + Table.Rows.Count.ToString)
                        Else
                            ht.Add(dataRiyobi.PropStrYoyakuDt, Table)
                        End If
                        ' 2015.12.25 UPD END↑ h.hagiwara
                    End If
                End If
            Next
            dataEXTC0102.PropHtFutaiDetail = ht
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

    ''' <summary>
    ''' 付帯情報取得(再計算)
    ''' </summary>
    ''' <param name="Adapter"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataEXTC0102"></param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>付帯情報取得するSQLの作成
    ''' <para>作成情報：2015/08/27 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function selectFutaiHeadderUp(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByRef Cn As NpgsqlConnection, _
                                       ByRef dataEXTC0102 As DataEXTC0102) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Try
            'SQL文(SELECT)
            Dim Cmd As New NpgsqlCommand
            Dim oldHt As New Hashtable
            Dim newHt As New Hashtable
            If dataEXTC0102.PropHtFutai IsNot Nothing Then
                oldHt = dataEXTC0102.PropHtFutai
            End If
            Cmd.Connection = Cn

            For Each dataRiyobi As CommonDataRiyobi In dataEXTC0102.PropListRiyobi
                If oldHt.ContainsKey(dataRiyobi.PropStrYoyakuDt) = False Then
                    Dim Table As New DataTable()
                    Cmd.CommandText = strEX05S097
                    Adapter.SelectCommand = Cmd
                    With Adapter.SelectCommand.Parameters
                        .Add(New NpgsqlParameter(":Riyobi", NpgsqlTypes.NpgsqlDbType.Varchar))
                        .Add(New NpgsqlParameter(":ShisetuKbn", NpgsqlTypes.NpgsqlDbType.Varchar))
                        .Add(New NpgsqlParameter(":SYANAI", NpgsqlTypes.NpgsqlDbType.Varchar))
                    End With
                    With Adapter.SelectCommand
                        .Parameters(0).Value = dataRiyobi.PropStrYoyakuDt
                        .Parameters(1).Value = dataEXTC0102.PropStrShisetuKbn
                        .Parameters(2).Value = dataEXTC0102.PropStrKashiKind
                    End With
                    Adapter.Fill(Table)
                    If Table.Rows.Count = 0 Then
                        '付帯設備初期情報が存在しない場合、仮行を設定する
                        Dim row As DataRow = Table.NewRow
                        row("yoyaku_no") = ""
                        row("yoyaku_dt") = dataRiyobi.PropStrYoyakuDt
                        row("total_fuzoku_kin") = 0
                        row("fuzoku_nm") = ""
                        row("tax_kin") = 0                                      ' 2015.11.13 ADD h.hagiwara
                        Table.Rows.Add(row)
                    End If
                    newHt.Add(dataRiyobi.PropStrYoyakuDt, Table)
                Else
                    newHt.Add(dataRiyobi.PropStrYoyakuDt, oldHt(dataRiyobi.PropStrYoyakuDt))
                End If
            Next
            dataEXTC0102.PropHtFutai = newHt

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

    ''' <summary>
    ''' 付帯情報取得(再計算)
    ''' </summary>
    ''' <param name="Adapter"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataEXTC0102"></param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>付帯情報取得するSQLの作成
    ''' <para>作成情報：2015/08/27 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function selectFutaiDetailUp(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByRef Cn As NpgsqlConnection, _
                                       ByRef dataEXTC0102 As DataEXTC0102) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Try
            'SQL文(SELECT)
            Dim oldHt As New Hashtable
            Dim newHt As New Hashtable
            Dim Cmd As New NpgsqlCommand
            oldHt = dataEXTC0102.PropHtFutaiDetail

            Cmd.Connection = Cn
            For Each dataRiyobi As CommonDataRiyobi In dataEXTC0102.PropListRiyobi
                If oldHt.ContainsKey(dataRiyobi.PropStrYoyakuDt) = False Then
                    Dim Table As New DataTable()
                    Cmd.CommandText = strEX05S096
                    Adapter.SelectCommand = Cmd
                    With Adapter.SelectCommand.Parameters
                        .Add(New NpgsqlParameter(":ShisetuKbn", NpgsqlTypes.NpgsqlDbType.Varchar))
                        .Add(New NpgsqlParameter(":Riyobi", NpgsqlTypes.NpgsqlDbType.Varchar))
                        .Add(New NpgsqlParameter(":SYANAI", NpgsqlTypes.NpgsqlDbType.Varchar))
                    End With
                    With Adapter.SelectCommand
                        .Parameters(0).Value = dataEXTC0102.PropStrShisetuKbn
                        .Parameters(1).Value = dataRiyobi.PropStrYoyakuDt
                        .Parameters(2).Value = dataEXTC0102.PropStrKashiKind
                    End With
                    Adapter.Fill(Table)
                    newHt.Add(dataRiyobi.PropStrYoyakuDt, Table)
                Else
                    newHt.Add(dataRiyobi.PropStrYoyakuDt, oldHt(dataRiyobi.PropStrYoyakuDt))
                End If
            Next
            dataEXTC0102.PropHtFutaiDetail = newHt
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

    ''' <summary>
    ''' メールアドレス取得
    ''' </summary>
    ''' <param name="Adapter"></param>
    ''' <param name="Cn"></param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>メールアドレス
    ''' <para>作成情報：2015/08/27 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function selecMailTo(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByRef Cn As NpgsqlConnection, _
                                       ByVal IsShonin As Boolean, ByVal Yoyakuno As String) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = String.Empty

        Try
            'SQL文(SELECT)
            'データアダプタに、SQLのSELECT文を設定
            strSQL = strEX05S099
            ' 2016.01.27 UPD START↓ h.hagiwara  
            'If IsShonin Then
            '    strSQL = strSQL + " AND shonin_flg = '1' "
            'Else                                                            ' 2015.11.20 ADD h.hagiwara
            '    strSQL = strSQL + " AND shonin_flg = '0' "                  ' 2015.11.20 ADD h.hagiwara
            'End If
            strSQL = strSQL + " AND shonin_flg = '1' "
            If IsShonin Then
            Else
                strSQL = strSQL + strEX05S0991
            End If
            ' 2016.01.27 UPD END↑ h.hagiwara  

            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)
            ' 2016.01.27 ADD START↓ h.hagiwara  
            Adapter.SelectCommand.Parameters.Add(New NpgsqlParameter(":YOYAKU_NO", NpgsqlTypes.NpgsqlDbType.Varchar))
            Adapter.SelectCommand.Parameters.Add(New NpgsqlParameter(":YOYAKU_NO2", NpgsqlTypes.NpgsqlDbType.Varchar))
            Adapter.SelectCommand.Parameters(0).Value = Yoyakuno
            Adapter.SelectCommand.Parameters(1).Value = Yoyakuno
            ' 2016.01.27 ADD END↑ h.hagiwara  
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

    ''' <summary>
    ''' 消費税取得
    ''' </summary>
    ''' <param name="Adapter"></param>
    ''' <param name="Cn"></param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>システムプロパティ
    ''' <para>作成情報：2015/08/28 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function selectTax(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByRef Cn As NpgsqlConnection, _
                                       ByVal Riyobi As String) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = String.Empty

        Try
            'SQL文(SELECT)
            'データアダプタに、SQLのSELECT文を設定
            strSQL = strEX05S095
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)
            Adapter.SelectCommand.Parameters.Add(New NpgsqlParameter(":Riyobi", NpgsqlTypes.NpgsqlDbType.Varchar))
            Adapter.SelectCommand.Parameters(0).Value = Riyobi
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

    ''' <summary>
    ''' EXASプロジェクト利用料情報取得
    ''' </summary>
    ''' <param name="Adapter"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataEXTC0102"></param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>請求情報を取得するSQLの作成
    ''' <para>作成情報：2015/09/03 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function selectExasRiyoryo(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByRef Cn As NpgsqlConnection, _
                                       ByRef dataEXTC0102 As DataEXTC0102) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Try
            'SQL文(SELECT)
            Dim ht As New Hashtable
            Dim Cmd As New NpgsqlCommand
            Cmd.Connection = Cn
            Cmd.CommandText = strEX38S004
            Adapter.SelectCommand = Cmd

            Dim dtBillPay As DataTable = dataEXTC0102.PropDsBillReq.Tables("BILLPAY_TBL")
            Dim index As Integer = 0
            Dim row As DataRow
            Do While index < dtBillPay.Rows.Count
                row = dtBillPay.Rows(index)
                '利用料を対象
                If row("seikyu_naiyo") = SEIKYU_NAIYOU_RIYO Or row("seikyu_naiyo") = SEIKYU_NAIYOU_RIYOFUTAI Then

                    With Adapter.SelectCommand.Parameters
                        .Add(New NpgsqlParameter(":YoyakuNo", NpgsqlTypes.NpgsqlDbType.Varchar))
                        .Add(New NpgsqlParameter(":SeikyuIraiNo", NpgsqlTypes.NpgsqlDbType.Varchar))
                    End With
                    With Adapter.SelectCommand
                        .Parameters(0).Value = dataEXTC0102.PropStrYoyakuNo
                        ' 2015.12.22 UPD START↓ h.hagiwara
                        '.Parameters(1).Value = row("seikyu_irai_no")
                        If row("seikyu_input_flg") = "" Or row("seikyu_input_flg") = "0" Then
                            .Parameters(1).Value = ""
                        Else
                            .Parameters(1).Value = row("seikyu_irai_no")
                        End If
                        ' 2015.12.22 UPD END↑ h.hagiwara
                    End With
                    Dim Table As New DataTable
                    Adapter.Fill(Table)
                    If Table.Rows.Count > 0 Then
                        ht.Add(row("seikyu_irai_no"), Table)
                    Else
                        Table = New DataTable
                        Dim baseTable As New DataTable
                        Dim TempTable As New DataTable
                        Dim beforeYm As String = ""
                        Dim strSetflg As String = ""                                              ' 2015.12.17 ADD h.hagiwara 金額表示不具合対応
                        Cmd.CommandText = strEX38S001 '初期データ(利用月分実行)
                        Adapter.SelectCommand = Cmd
                        Adapter.Fill(baseTable)
                        If baseTable.Rows.Count = 0 Then
                            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, "利用料の科目を科目マスタに設定して下さい。", Nothing, Nothing)
                        End If
                        Table = baseTable.Clone
                        ' 2015.12.17 UPD START↓ h.hagiwara 金額表示不具合対応
                        'For Each dataRiyobi As CommonDataRiyobi In dataEXTC0102.PropListRiyobi
                        '    Dim dt As Date = dataRiyobi.PropStrYoyakuDt
                        '    Dim currentYm As String = dt.ToString("yyyyMM")
                        '    Dim tempRow As DataRow = baseTable.Rows(0)
                        '    If beforeYm = currentYm = False Then
                        '        'Row作成
                        '        tempRow("tekiyo1") = row("seikyu_title1")
                        '        tempRow("tekiyo2") = row("seikyu_title2")
                        '        tempRow("riyo_ym") = currentYm
                        '        tempRow("keijo_kin") = dataRiyobi.PropIntRiyoKin
                        '        tempRow("tax_kin") = Integer.Parse(Math.Round((dataRiyobi.PropIntRiyoKin * dataEXTC0102.propDblTax), MidpointRounding.AwayFromZero))
                        '        selecContents(Adapter, Cn, dataRiyobi.PropStrYoyakuDt, tempRow)
                        '        Table.ImportRow(tempRow)
                        '        beforeYm = currentYm
                        '    Else
                        '        tempRow = Table.Rows(Table.Rows.Count - 1)
                        '        tempRow("keijo_kin") = tempRow("keijo_kin") + dataRiyobi.PropIntRiyoKin
                        '        tempRow("tax_kin") = tempRow("tax_kin") + Integer.Parse(Math.Round((dataRiyobi.PropIntRiyoKin * dataEXTC0102.propDblTax), MidpointRounding.AwayFromZero))
                        '    End If
                        'Next
                        For Each dataRiyobi As CommonDataRiyobi In dataEXTC0102.PropListRiyobi
                            Dim dt As Date = dataRiyobi.PropStrYoyakuDt
                            Dim currentYm As String = dt.ToString("yyyyMM")
                            Dim tempRow As DataRow = baseTable.Rows(0)
                            If beforeYm = currentYm = False Then
                                'Row作成
                                tempRow("tekiyo1") = row("seikyu_title1")
                                tempRow("tekiyo2") = row("seikyu_title2")
                                tempRow("riyo_ym") = currentYm
                                If strSetflg = "" Then
                                    tempRow("keijo_kin") = row("kakutei_kin")
                                    tempRow("tax_kin") = row("tax_kin")
                                    strSetflg = "1"
                                Else
                                    tempRow("keijo_kin") = "0"
                                    tempRow("tax_kin") = "0"
                                End If
                                selecContents(Adapter, Cn, dataRiyobi.PropStrYoyakuDt, tempRow)
                                Table.ImportRow(tempRow)
                                beforeYm = currentYm
                            End If
                        Next
                        ' 2015.12.17 UPD END↑ h.hagiwara 金額表示不具合対応
                        'ソート処理
                        Dim srtTable As New DataTable()
                        If Table.Rows.Count > 0 Then
                            srtTable = Table.Clone
                            Dim rows As DataRow() = Table.Select(Nothing, "riyo_ym ASC, kamoku_nm ASC")
                            For Each srtRow As DataRow In rows
                                srtTable.ImportRow(srtRow)
                            Next
                            Table = srtTable
                        End If
                        If IsDBNull(row("seikyu_irai_no")) = False Then
                            ht.Add(row("seikyu_irai_no"), Table)
                        Else
                            ht.Add(index.ToString, Table)
                        End If
                        Adapter.SelectCommand = New NpgsqlCommand(strEX38S004, Cn)
                    End If
                End If
                index = index + 1
            Loop
            dataEXTC0102.PropHtExasPro = ht
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

    ''' <summary>
    ''' EXASプロジェクト利用料情報取得
    ''' </summary>
    ''' <param name="Adapter"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataEXTC0102"></param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>請求情報を取得するSQLの作成
    ''' <para>作成情報：2015/09/03 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function selectExasRiyoryoUp(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByRef Cn As NpgsqlConnection, _
                                       ByRef dataEXTC0102 As DataEXTC0102) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Try
            'SQL文(SELECT)
            Dim oldHt As Hashtable = dataEXTC0102.PropHtExasPro
            Dim newHt As New Hashtable
            Dim Cmd As New NpgsqlCommand
            Cmd.Connection = Cn
            Cmd.CommandText = strEX38S001
            Adapter.SelectCommand = Cmd

            Dim dtBillPay As DataTable = dataEXTC0102.PropDsBillReq.Tables("BILLPAY_TBL")
            Dim index As Integer = 0
            Dim row As DataRow
            Dim htKey As String
            Dim newTable As New DataTable
            Dim tempTable As New DataTable
            'Dim intRiyoKeijo() As Integer = {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}                    ' 2015.12.21 UPD h.hagiwara
            Dim intRiyoKeijo() As Long = {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}                        ' 2015.12.21 UPD h.hagiwara
            '請求入力済の合計値確認
            Do While index < dtBillPay.Rows.Count
                row = dtBillPay.Rows(index)
                If row("seikyu_naiyo") = SEIKYU_NAIYOU_RIYO Or row("seikyu_naiyo") = SEIKYU_NAIYOU_RIYOFUTAI Then
                    If row("seikyu_input_flg") = "1" Then
                        Dim prjIndex As Integer = 0
                        Dim Table As New DataTable
                        Dim prjRow As DataRow
                        If IsDBNull(row("seikyu_irai_no")) Then
                            htKey = index.ToString
                        Else
                            htKey = row("seikyu_irai_no")
                        End If
                        Table = oldHt(htKey)
                        Do While prjIndex < Table.Rows.Count
                            prjRow = Table.Rows(prjIndex)
                            intRiyoKeijo(prjIndex) += prjRow("keijo_kin")
                            prjIndex += 1
                        Loop
                    End If
                End If
                index += 1
            Loop
            index = 0
            Do While index < dtBillPay.Rows.Count
                row = dtBillPay.Rows(index)
                '利用料を対象
                If row("seikyu_naiyo") = SEIKYU_NAIYOU_RIYO Or row("seikyu_naiyo") = SEIKYU_NAIYOU_RIYOFUTAI Then
                    Dim Table As New DataTable
                    If IsDBNull(row("seikyu_irai_no")) Then
                        htKey = index.ToString
                    Else
                        htKey = row("seikyu_irai_no")
                    End If
                    Table = oldHt(htKey)
                    If row("seikyu_input_flg") = "1" = False Then
                        Dim beforeYm As String = ""
                        Adapter.Fill(tempTable)
                        newTable = tempTable.Clone
                        For Each dataRiyobi As CommonDataRiyobi In dataEXTC0102.PropListRiyobi
                            Dim dt As Date = dataRiyobi.PropStrYoyakuDt
                            Dim currentYm As String = dt.ToString("yyyyMM")
                            Dim tempRow As DataRow = tempTable.Rows(0)
                            Dim currentRows As DataRow() = Table.Select(String.Format("riyo_ym = {0}", currentYm))
                            Dim currentRow As DataRow
                            If currentRows.Count = 0 Then
                                'Row作成
                                tempRow("tekiyo1") = row("seikyu_title1")
                                tempRow("tekiyo2") = row("seikyu_title2")
                                tempRow("riyo_ym") = currentYm
                                tempRow("keijo_kin") = dataRiyobi.PropIntRiyoKin
                                'tempRow("tax_kin") = Integer.Parse(Math.Round((dataRiyobi.PropIntRiyoKin * dataEXTC0102.propDblTax), MidpointRounding.AwayFromZero))               ' 2015.12.21 UPD h.hagiwara
                                tempRow("tax_kin") = Long.Parse(Math.Round((dataRiyobi.PropIntRiyoKin * dataEXTC0102.propDblTax), MidpointRounding.AwayFromZero))                   ' 2015.12.21 UPD h.hagiwara
                                selecContents(Adapter, Cn, dataRiyobi.PropStrYoyakuDt, tempRow)
                                newTable.ImportRow(tempRow)
                                beforeYm = currentYm
                            Else
                                currentRow = currentRows(0)
                                If beforeYm = currentYm = False Then
                                    currentRow("keijo_kin") = dataRiyobi.PropIntRiyoKin
                                    'currentRow("tax_kin") = Integer.Parse(Math.Round((dataRiyobi.PropIntRiyoKin * dataEXTC0102.propDblTax), MidpointRounding.AwayFromZero))        ' 2015.12.21 UPD h.hagiwara
                                    currentRow("tax_kin") = Long.Parse(Math.Round((dataRiyobi.PropIntRiyoKin * dataEXTC0102.propDblTax), MidpointRounding.AwayFromZero))            ' 2015.12.21 UPD h.hagiwara
                                    beforeYm = currentYm
                                    newTable.ImportRow(currentRow)
                                Else
                                    currentRow = newTable.Select(String.Format("riyo_ym = {0}", currentYm))(0)
                                    currentRow("keijo_kin") = currentRow("keijo_kin") + dataRiyobi.PropIntRiyoKin
                                    'currentRow("tax_kin") = Integer.Parse(Math.Round((currentRow("keijo_kin") * dataEXTC0102.propDblTax), MidpointRounding.AwayFromZero))          ' 2015.12.21 UPD h.hagiwara
                                    currentRow("tax_kin") = Long.Parse(Math.Round((currentRow("keijo_kin") * dataEXTC0102.propDblTax), MidpointRounding.AwayFromZero))              ' 2015.12.21 UPD h.hagiwara
                                    beforeYm = currentYm
                                End If
                            End If
                        Next
                        'ソート処理
                        Dim srtTable As New DataTable()
                        If newTable.Rows.Count > 0 Then
                            srtTable = newTable.Clone
                            Dim rows As DataRow() = newTable.Select(Nothing, "riyo_ym ASC, kamoku_nm ASC")
                            For Each srtRow As DataRow In rows
                                srtTable.ImportRow(srtRow)
                            Next
                            newTable = srtTable
                        End If
                        '請求金額調整
                        Dim prjIndex As Integer = 0
                        Dim prjRow As DataRow
                        Do While prjIndex < newTable.Rows.Count
                            prjRow = newTable.Rows(prjIndex)
                            prjRow("keijo_kin") -= intRiyoKeijo(prjIndex)
                            'prjRow("tax_kin") = Integer.Parse(Math.Round((prjRow("keijo_kin") * dataEXTC0102.propDblTax), MidpointRounding.AwayFromZero))              ' 2015.12.21 UPD h.hagiwara
                            prjRow("tax_kin") = Long.Parse(Math.Round((prjRow("keijo_kin") * dataEXTC0102.propDblTax), MidpointRounding.AwayFromZero))                  ' 2015.12.21 UPD h.hagiwara
                            intRiyoKeijo(prjIndex) += prjRow("keijo_kin")
                            prjIndex += 1
                        Loop
                    Else
                        '追加分の作成
                        newTable = Table.Copy()
                    End If
                    newHt.Add(htKey, newTable)
                End If
                index = index + 1
            Loop
            dataEXTC0102.PropHtExasPro = newHt
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

    ''' <summary>
    ''' EXASプロジェクト付帯情報取得
    ''' </summary>
    ''' <param name="Adapter"></param>
    ''' <param name="Cn"></param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>請求情報を取得するSQLの作成
    ''' <para>作成情報：2015/09/03 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function selectExasFutai(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByRef Cn As NpgsqlConnection, _
                                       ByRef dataEXTC0102 As DataEXTC0102) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Try
            'SQL文(SELECT)
            Dim ht As New Hashtable
            Dim Cmd As New NpgsqlCommand
            Cmd.Connection = Cn
            Cmd.CommandText = strEX38S005
            Adapter.SelectCommand = Cmd

            Dim dtBillPay As DataTable = dataEXTC0102.PropDsBillReq.Tables("BILLPAY_TBL")
            Dim index As Integer = 0
            Dim row As DataRow
            Do While index < dtBillPay.Rows.Count
                row = dtBillPay.Rows(index)
                '付帯料を対象
                If row("seikyu_naiyo") = SEIKYU_NAIYOU_FUTAI Or row("seikyu_naiyo") = SEIKYU_NAIYOU_RIYOFUTAI Then
                    With Adapter.SelectCommand.Parameters
                        .Add(New NpgsqlParameter(":YoyakuNo", NpgsqlTypes.NpgsqlDbType.Varchar))
                        .Add(New NpgsqlParameter(":SeikyuIraiNo", NpgsqlTypes.NpgsqlDbType.Varchar))
                        .Add(New NpgsqlParameter(":ShisetuKbn", NpgsqlTypes.NpgsqlDbType.Varchar))
                        .Add(New NpgsqlParameter(":Riyobi", NpgsqlTypes.NpgsqlDbType.Varchar))
                    End With
                    With Adapter.SelectCommand
                        .Parameters(0).Value = dataEXTC0102.PropStrYoyakuNo
                        ' 2015.12.22 UPD START↓ h.hagiwara
                        '.Parameters(1).Value = row("seikyu_irai_no")
                        If row("seikyu_input_flg") = "" Or row("seikyu_input_flg") = "0" Then
                            .Parameters(1).Value = ""
                        Else
                            .Parameters(1).Value = row("seikyu_irai_no")
                        End If
                        ' 2015.12.22 UPD END↑ h.hagiwara
                        .Parameters(2).Value = dataEXTC0102.PropStrShisetuKbn
                        .Parameters(3).Value = ""
                    End With
                    Dim Table As New DataTable()
                    Adapter.Fill(Table)
                    If Table.Rows.Count > 0 Then

                        ' --- 2019/08/07 軽減税率対応 Start E.Okuda@Compass ---
                        calExasFutai2(dataEXTC0102.PropHtFutaiDetail, Table)
                        ' --- 2019/08/07 軽減税率対応 End E.Okuda@Compass ---

                        ht.Add(row("seikyu_irai_no"), Table)
                    Else
                        calcExasFutai(Adapter, Cn, dataEXTC0102.PropHtFutaiDetail, Table, dataEXTC0102.propDblTax, dataEXTC0102.PropListRiyobi, row)
                        If IsDBNull(row("seikyu_irai_no")) = False Then
                            ht.Add(row("seikyu_irai_no"), Table)
                        Else
                            ht.Add(index.ToString, Table)
                        End If
                        Adapter.SelectCommand = New NpgsqlCommand(strEX38S005, Cn)
                    End If
                End If
                index = index + 1
            Loop
            dataEXTC0102.PropHtExasProFutai = ht
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常終了
            Return True
        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
            '例外処理
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Exasプロジェクト明細(付帯)の設定を行う
    ''' </summary>
    ''' <param name="Adapter"></param>
    ''' <param name="Cn"></param>
    ''' <param name="htFutaiDetail"></param>
    ''' <param name="exasFutaitable"></param>
    ''' <param name="dblTax"></param>
    ''' <param name="lstRiyobi"></param>
    ''' <param name="billPayRow"></param>
    ''' <remarks></remarks>
    Public Sub calcExasFutai(ByRef Adapter As NpgsqlDataAdapter, _
                             ByRef Cn As NpgsqlConnection, _
                             ByRef htFutaiDetail As Hashtable, _
                             ByRef exasFutaitable As DataTable, _
                             ByVal dblTax As Double, _
                             ByVal lstRiyobi As ArrayList, _
                             ByVal billPayRow As DataRow)
        Try
            Dim dt As Date
            Dim currentYm As String = ""
            'グループキーで付帯計上の計算
            Dim row As DataRow
            Dim newRow As DataRow
            Dim htTempExasFutai As New Hashtable
            Dim tableFutaiDetail As DataTable
            For Each Key As String In htFutaiDetail.Keys
                tableFutaiDetail = htFutaiDetail(Key)
                dt = Key
                currentYm = dt.ToString("yyyyMM")
                Dim futaiIndex As Integer = 0
                Do While futaiIndex < tableFutaiDetail.Rows.Count
                    row = tableFutaiDetail.Rows(futaiIndex)
                    If htTempExasFutai.ContainsKey(currentYm + row("shukei_grp")) Then
                        newRow = htTempExasFutai(currentYm + row("shukei_grp"))
                        newRow("keijo_kin") = newRow("keijo_kin") + row("futai_shokei")
                        'newRow("tax_kin") = newRow("tax_kin") + Integer.Parse(Math.Round((row("futai_shokei") * dblTax), MidpointRounding.AwayFromZero))  ' 2015/11/13 UPD h.hagiwara
                        newRow("tax_kin") = newRow("tax_kin") + row("tax_kin")                                                                             ' 2015/11/13 UPD h.hagiwara
                        htTempExasFutai.Remove(currentYm + row("shukei_grp"))
                    Else
                        newRow = exasFutaitable.NewRow
                        newRow("riyo_ym") = currentYm
                        newRow("shukei_grp") = row("shukei_grp")
                        newRow("kamoku_nm") = row("kamoku_nm")
                        newRow("saimoku_nm") = row("saimoku_nm")
                        newRow("uchi_nm") = row("uchi_nm")
                        newRow("shosai_nm") = row("shosai_nm")
                        newRow("keijo_kin") = row("futai_shokei")
                        'newRow("tax_kin") = Integer.Parse(Math.Round((row("futai_shokei") * dblTax), MidpointRounding.AwayFromZero))   ' 2015/11/13 UPD h.hagiwara
                        newRow("tax_kin") = row("tax_kin")                                                                              ' 2015/11/13 UPD h.hagiwara
                        newRow("kamoku_cd") = row("kamoku_cd")
                        newRow("saimoku_cd") = row("saimoku_cd")
                        newRow("uchi_cd") = row("uchi_cd")
                        newRow("shosai_cd") = row("shosai_cd")
                        newRow("seikyu_naiyo") = SEIKYU_NAIYOU_FUTAI
                        newRow("tekiyo1") = billPayRow("seikyu_title1")
                        newRow("tekiyo2") = billPayRow("seikyu_title2")
                        newRow("karikamoku_cd") = row("karikamoku_cd")
                        newRow("kari_saimoku_cd") = row("kari_saimoku_cd")
                        newRow("kari_uchi_cd") = row("kari_uchi_cd")
                        newRow("kari_shosai_cd") = row("kari_shosai_cd")
                        ' --- 2019/08/07 軽減税率対応 Start E.Okuda@Compass ---
                        newRow("zeiritsu") = row("zeiritsu")
                        newRow("notax_flg") = row("notax_flg")
                        ' --- 2019/08/07 軽減税率対応 End E.Okuda@Compass ---

                        'プロジェクトを検索する
                        selecContents(Adapter, Cn, Key, newRow)
                    End If
                    Debug.WriteLine(Key + row("shukei_grp"))
                    htTempExasFutai.Add(currentYm + row("shukei_grp"), newRow)
                    futaiIndex = futaiIndex + 1
                Loop
            Next
            'Exas付帯情報の格納
            Dim futaiRow As DataRow
            For Each Key As String In htTempExasFutai.Keys
                futaiRow = htTempExasFutai(Key)
                Dim exasIndex As Integer = 0
                Dim exasRow As DataRow
                Dim blnExist As Boolean = False
                Dim lineIndex As New Integer
                Do While exasIndex < exasFutaitable.Rows.Count
                    exasRow = exasFutaitable(exasIndex)
                    If exasRow("riyo_ym") = futaiRow("riyo_ym") And exasRow("shukei_grp") = futaiRow("shukei_grp") Then
                        blnExist = True
                        lineIndex = exasIndex
                    End If
                    exasIndex = exasIndex + 1
                Loop
                If blnExist = False Then
                    exasFutaitable.Rows.Add(futaiRow)
                Else
                    exasFutaitable.Rows.RemoveAt(lineIndex)
                    exasFutaitable.Rows.Add(futaiRow)
                End If
            Next
            'ソート処理
            Dim srtExasFutaiTable As New DataTable()
            If exasFutaitable.Rows.Count > 0 Then
                srtExasFutaiTable = exasFutaitable.Clone
                Dim rows As DataRow() = exasFutaitable.Select(Nothing, "riyo_ym ASC, shukei_grp ASC")
                For Each srtRow As DataRow In rows
                    For Each dataRiyobi As CommonDataRiyobi In lstRiyobi
                        Dim tempDt As Date = dataRiyobi.PropStrYoyakuDt
                        If tempDt.ToString("yyyyMM") = srtRow("riyo_ym") Then
                            srtExasFutaiTable.ImportRow(srtRow)
                            Exit For
                        End If
                    Next
                Next
                exasFutaitable = srtExasFutaiTable
            End If
        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
            '例外処理
            Throw ex
        End Try
    End Sub


    ' --- 2019/08/07 軽減税率対応 Start E.Okuda@Compass ---

    ''' <summary>
    '''  Exasプロジェクト明細(付帯)に「税率」、「税区分」設定を行う
    ''' </summary>
    ''' <param name="htFutaiDetail"></param>
    ''' <param name="exasFutaitable"></param>
    ''' <remarks>
    ''' <para>作成情報：2019/08/07 E.Okuda@Compass
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Sub calExasFutai2(ByRef htFutaiDetail As Hashtable, ByRef exasFutaitable As DataTable)
        Try
            Dim dt As Date
            Dim currentYm As String = ""
            Dim htTempExasFutai As New Hashtable
            Dim tableFutaiDetail As DataTable
            Dim futaiRow As DataRow
            Dim exasRow As DataRow
            Dim futaiIndex As Integer = 0
            Dim exasIndex As Integer = 0

            Do While exasIndex < exasFutaitable.Rows.Count
                exasRow = exasFutaitable(exasIndex)
                For Each Key As String In htFutaiDetail.Keys
                    tableFutaiDetail = htFutaiDetail(Key)
                    dt = Key
                    currentYm = dt.ToString("yyyyMM")

                    Do While futaiIndex < tableFutaiDetail.Rows.Count
                        futaiRow = tableFutaiDetail.Rows(futaiIndex)

                        If exasRow("riyo_ym") = currentYm And exasRow("shukei_grp") = futaiRow("shukei_grp") Then
                            exasRow("zeiritsu") = futaiRow("zeiritsu")
                            exasRow("notax_flg") = futaiRow("notax_flg")

                            Exit Do
                        End If
                        futaiIndex = futaiIndex + 1
                    Loop
                Next

                exasIndex = exasIndex + 1
            Loop

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            '例外処理
            Throw ex
        End Try
    End Sub
    ' --- 2019/08/07 軽減税率対応 End E.Okuda@Compass ---


    ''' <summary>
    ''' EXASプロジェクト付帯情報取得
    ''' </summary>
    ''' <param name="Adapter"></param>
    ''' <param name="Cn"></param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>請求情報を取得するSQLの作成
    ''' <para>作成情報：2015/09/03 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function selectExasFutaiUp(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByRef Cn As NpgsqlConnection, _
                                       ByRef dataEXTC0102 As DataEXTC0102) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Try
            'SQL文(SELECT)
            Dim ht As Hashtable = dataEXTC0102.PropHtExasProFutai
            Dim Cmd As New NpgsqlCommand
            Cmd.Connection = Cn
            Cmd.CommandText = strEX38S005
            Adapter.SelectCommand = Cmd

            Dim dtBillPay As DataTable = dataEXTC0102.PropDsBillReq.Tables("BILLPAY_TBL")
            Dim index As Integer = 0
            Dim row As DataRow
            Dim htKey As String
            'Dim intFutaiKeijo() As Integer = {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}                  ' 2015.12.21 UPD h.hagiwara
            Dim intFutaiKeijo() As Long = {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}                      ' 2015.12.21 UPD h.hagiwara
            '請求入力済の合計値確認
            Do While index < dtBillPay.Rows.Count
                row = dtBillPay.Rows(index)
                If row("seikyu_naiyo") = SEIKYU_NAIYOU_FUTAI Or row("seikyu_naiyo") = SEIKYU_NAIYOU_RIYOFUTAI Then
                    If row("seikyu_input_flg") = "1" Then
                        Dim prjIndex As Integer = 0
                        Dim Table As New DataTable
                        Dim prjRow As DataRow
                        If IsDBNull(row("seikyu_irai_no")) Then
                            htKey = index.ToString
                        Else
                            htKey = row("seikyu_irai_no")
                        End If
                        Table = ht(htKey)
                        Do While prjIndex < Table.Rows.Count
                            prjRow = Table.Rows(prjIndex)
                            If IsDBNull(prjRow("keijo_kin")) Then
                                intFutaiKeijo(prjIndex) += 0
                            Else
                                intFutaiKeijo(prjIndex) += prjRow("keijo_kin")
                            End If
                            prjIndex += 1
                        Loop
                    End If
                End If
                index += 1
            Loop
            index = 0
            Do While index < dtBillPay.Rows.Count
                row = dtBillPay.Rows(index)
                Dim Table As New DataTable
                If IsDBNull(row("seikyu_irai_no")) Then
                    htKey = index.ToString
                Else
                    htKey = row("seikyu_irai_no")
                End If
                Table = ht(htKey)
                '付帯料を対象
                If row("seikyu_input_flg") = "1" = False Then
                    If row("seikyu_naiyo") = SEIKYU_NAIYOU_FUTAI Or row("seikyu_naiyo") = SEIKYU_NAIYOU_RIYOFUTAI Then
                        With Adapter.SelectCommand.Parameters
                            .Add(New NpgsqlParameter(":YoyakuNo", NpgsqlTypes.NpgsqlDbType.Varchar))
                            .Add(New NpgsqlParameter(":SeikyuIraiNo", NpgsqlTypes.NpgsqlDbType.Varchar))
                            .Add(New NpgsqlParameter(":ShisetuKbn", NpgsqlTypes.NpgsqlDbType.Varchar))
                            .Add(New NpgsqlParameter(":Riyobi", NpgsqlTypes.NpgsqlDbType.Varchar))
                        End With
                        With Adapter.SelectCommand
                            .Parameters(0).Value = dataEXTC0102.PropStrYoyakuNo
                            .Parameters(1).Value = row("seikyu_irai_no")
                            .Parameters(2).Value = dataEXTC0102.PropStrShisetuKbn
                            .Parameters(3).Value = ""
                        End With
                        calcExasFutai(Adapter, Cn, dataEXTC0102.PropHtFutaiDetail, Table, dataEXTC0102.propDblTax, dataEXTC0102.PropListRiyobi, row)
                        '請求金額調整
                        Dim prjIndex As Integer = 0
                        Dim prjRow As DataRow
                        Do While prjIndex < Table.Rows.Count
                            prjRow = Table.Rows(prjIndex)
                            prjRow("keijo_kin") -= intFutaiKeijo(prjIndex)
                            'prjRow("tax_kin") = Integer.Parse(Math.Round((prjRow("keijo_kin") * dataEXTC0102.propDblTax), MidpointRounding.AwayFromZero))   ' 2015.11.13 DEL h.hagiwara
                            If dataEXTC0102.PropStrKashiKind = CommonDeclareEXT.KASHI_KIND_HOUSE Then
                                prjRow("tax_kin") = 0
                            Else
                                prjRow("tax_kin") = prjRow("tax_kin")
                            End If
                            intFutaiKeijo(prjIndex) += prjRow("keijo_kin")
                            prjIndex += 1
                        Loop
                        ht(htKey) = Table
                    End If
                End If
                index = index + 1
            Loop
            dataEXTC0102.PropHtExasProFutai = ht
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常終了
            Return True
        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
            '例外処理
            Return False
        End Try
    End Function

    ''' <summary>
    ''' プロジェクト(コンテンツ)情報取得
    ''' </summary>
    ''' <param name="Adapter"></param>
    ''' <param name="Cn"></param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>メールアドレス
    ''' <para>作成情報：2015/08/27 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function selecContents(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByRef Cn As NpgsqlConnection, _
                                       ByVal StrRiyobi As String, _
                                       ByRef exasRow As DataRow) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = String.Empty

        Try
            'SQL文(SELECT)
            'データアダプタに、SQLのSELECT文を設定
            Adapter.SelectCommand = New NpgsqlCommand(strEX41S006, Cn)
            Adapter.SelectCommand.Parameters.Add(New NpgsqlParameter(":ContentUchiNm", NpgsqlTypes.NpgsqlDbType.Varchar))
            Adapter.SelectCommand.Parameters.Add(New NpgsqlParameter(":Riyobi", NpgsqlTypes.NpgsqlDbType.Varchar))
            'Adapter.SelectCommand.Parameters(0).Value = "貸館" + commonLogicEXT.convYmDateStr(StrRiyobi)                                               ' 2015.12.10 UPD h.hagiwara
            Adapter.SelectCommand.Parameters(0).Value = "リハスタ" + StrConv(commonLogicEXT.convYmDateStr2(StrRiyobi), VbStrConv.Wide)                  ' 2015.12.10 UPD h.hagiwara
            Adapter.SelectCommand.Parameters(1).Value = StrRiyobi.Replace("/", "")
            Dim table As New DataTable
            Adapter.Fill(table)
            If table.Rows.Count > 0 Then
                exasRow("event_nm") = table.Rows(0)("event_nm")
                exasRow("content_uchi_nm") = table.Rows(0)("content_uchi_nm")
                exasRow("content_cd") = table.Rows(0)("event_cd")
                exasRow("content_uchi_cd") = table.Rows(0)("content_uchi_cd")
            End If

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Adapter.SelectCommand)

            '正常終了
            Return True
        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            '例外処理
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 付帯表削除(予約番号、notin予約日)
    ''' </summary>
    ''' <param name="Cmd"></param>
    ''' <param name="Cn"></param>
    ''' <param name="YoyakuNo"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' <para>作成情報：2015/09/17 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function deleteFutai(ByRef Cmd As NpgsqlCommand, _
                                   ByRef Cn As NpgsqlConnection, _
                                   ByVal YoyakuNo As String, _
                                   ByVal LstRiyobi As ArrayList
                                   ) As Boolean
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = String.Empty
        Dim strWhere As String = "("
        Try
            strSQL = strEX05D005
            Dim isFirst As Boolean = True
            For Each dataRiyobi As CommonDataRiyobi In LstRiyobi
                If isFirst = False Then
                    strWhere = strWhere + ","
                Else
                    isFirst = False
                End If
                strWhere = strWhere + "'" + dataRiyobi.PropStrYoyakuDt + "'"
            Next
            If isFirst = True Then
                Return True
            End If
            strWhere = strWhere + ")"
            strSQL = strSQL + strWhere
            Cmd = New NpgsqlCommand(strSQL, Cn)

            '値を設定
            Cmd.Parameters.Add(New NpgsqlParameter(":YoyakuNo", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters(0).Value = YoyakuNo

            Cmd.ExecuteScalar()

            '正常終了
            Return True
        Catch ex As Exception
            '例外処理
            puErrMsg = ex.Message
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            Return False
        End Try

    End Function

    ''' <summary>
    ''' 付帯テーブルの登録
    ''' </summary>
    ''' <param name="Cmd"></param>
    ''' <param name="Cn"></param>
    ''' <param name="FutaiRow"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' <para>作成情報：2015/08/05 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function registerFutai(ByRef Cmd As NpgsqlCommand, _
                                   ByRef Cn As NpgsqlConnection, _
                                   ByRef FutaiRow As DataRow, _
                                   ByVal YoyakuNo As String) As Boolean

        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = String.Empty
        Dim strNm As String = ""                ' 2015.11.27 ADD h.hagiwara

        Try
            'SQL文を設定
            If String.IsNullOrEmpty(FutaiRow("yoyaku_no")) Then
                'insert
                FutaiRow("yoyaku_no") = YoyakuNo
                strSQL = strEX05I005
            Else
                'update
                strSQL = strEX05U005
            End If

            Cmd = New NpgsqlCommand(strSQL, Cn)

            '値を設定
            Cmd.Parameters.Add(New NpgsqlParameter(":YoyakuNo", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":Riyobi", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":FuzokuNm", NpgsqlTypes.NpgsqlDbType.Varchar))
            'Cmd.Parameters.Add(New NpgsqlParameter(":TotalFuzokuKin", NpgsqlTypes.NpgsqlDbType.Integer))                ' 2015.12.21 UPD h.hagiwara
            Cmd.Parameters.Add(New NpgsqlParameter(":TotalFuzokuKin", NpgsqlTypes.NpgsqlDbType.Bigint))                  ' 2015.12.21 UPD h.hagiwara
            Cmd.Parameters.Add(New NpgsqlParameter(":AddUserCd", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":UpUserCd", NpgsqlTypes.NpgsqlDbType.Varchar))

            Cmd.Parameters(0).Value = YoyakuNo
            Cmd.Parameters(1).Value = FutaiRow("yoyaku_dt")
            ' 2015.11.27 START↓ UPD h.hagiwara 文字数オーバー対応
            'Cmd.Parameters(2).Value = FutaiRow("fuzoku_nm")
            strNm = FutaiRow("fuzoku_nm")
            If strNm.Length > 300 Then
                Cmd.Parameters(2).Value = strNm.Substring(0, 300)
            Else
                Cmd.Parameters(2).Value = strNm
            End If
            ' 2015.11.27 END↑ UPD h.hagiwara 文字数オーバー対応
            Cmd.Parameters(3).Value = FutaiRow("total_fuzoku_kin")
            Cmd.Parameters(4).Value = CommonEXT.PropComStrUserId
            Cmd.Parameters(5).Value = CommonEXT.PropComStrUserId

            Cmd.ExecuteScalar()
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常終了
            Return True
        Catch ex As Exception
            '例外処理
            puErrMsg = ex.Message
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Cmd)
            Return False
        End Try

    End Function

    ''' <summary>
    ''' 付帯明細表削除(予約番号)
    ''' </summary>
    ''' <param name="Cmd"></param>
    ''' <param name="Cn"></param>
    ''' <param name="YoyakuNo"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' <para>作成情報：2015/09/17 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function deleteFutaiDetail(ByRef Cmd As NpgsqlCommand, _
                                   ByRef Cn As NpgsqlConnection, _
                                   ByVal YoyakuNo As String
                                   ) As Boolean
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = String.Empty
        Try
            strSQL = strEX05D005D
            Cmd = New NpgsqlCommand(strSQL, Cn)

            '値を設定
            Cmd.Parameters.Add(New NpgsqlParameter(":YoyakuNo", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters(0).Value = YoyakuNo

            Cmd.ExecuteScalar()

            '正常終了
            Return True
        Catch ex As Exception
            '例外処理
            puErrMsg = ex.Message
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            Return False
        End Try

    End Function

    ''' <summary>
    ''' 付帯明細テーブルの登録
    ''' </summary>
    ''' <param name="Cmd"></param>
    ''' <param name="Cn"></param>
    ''' <param name="FutaiDetailRow"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' <para>作成情報：2015/08/05 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function registerFutaiDetail(ByRef Cmd As NpgsqlCommand, _
                                   ByRef Cn As NpgsqlConnection, _
                                   ByRef FutaiDetailRow As DataRow, _
                                   ByVal YoyakuNo As String) As Boolean

        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = String.Empty
        Try
            strSQL = strEX05I005D
            Cmd = New NpgsqlCommand(strSQL, Cn)

            '値を設定
            Cmd.Parameters.Add(New NpgsqlParameter(":YoyakuNo", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":Riyobi", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":FutaiBunruiCd", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":FutaiCd", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":FutaiTanka", NpgsqlTypes.NpgsqlDbType.Integer))
            Cmd.Parameters.Add(New NpgsqlParameter(":FutaiSu", NpgsqlTypes.NpgsqlDbType.Integer))
            'Cmd.Parameters.Add(New NpgsqlParameter(":FutaiShokei", NpgsqlTypes.NpgsqlDbType.Integer))                          ' 2015.12.21 UPD h.hagiwara
            Cmd.Parameters.Add(New NpgsqlParameter(":FutaiShokei", NpgsqlTypes.NpgsqlDbType.Bigint))                            ' 2015.12.21 UPD h.hagiwara
            Cmd.Parameters.Add(New NpgsqlParameter(":FutaiChosei", NpgsqlTypes.NpgsqlDbType.Integer))
            'Cmd.Parameters.Add(New NpgsqlParameter(":FutaiKin", NpgsqlTypes.NpgsqlDbType.Integer))                             ' 2015.12.21 UPD h.hagiwara
            Cmd.Parameters.Add(New NpgsqlParameter(":FutaiKin", NpgsqlTypes.NpgsqlDbType.Bigint))                               ' 2015.12.21 UPD h.hagiwara
            Cmd.Parameters.Add(New NpgsqlParameter(":FutaiBiko", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":AddUserCd", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":UpUserCd", NpgsqlTypes.NpgsqlDbType.Varchar))

            Cmd.Parameters(0).Value = YoyakuNo
            Cmd.Parameters(1).Value = FutaiDetailRow("yoyaku_dt")
            Cmd.Parameters(2).Value = FutaiDetailRow("futai_bunrui_cd")
            Cmd.Parameters(3).Value = FutaiDetailRow("futai_cd")
            Cmd.Parameters(4).Value = FutaiDetailRow("futai_tanka")
            Cmd.Parameters(5).Value = FutaiDetailRow("futai_su")
            Cmd.Parameters(6).Value = FutaiDetailRow("futai_shokei")
            Cmd.Parameters(7).Value = FutaiDetailRow("futai_chosei")
            Cmd.Parameters(8).Value = FutaiDetailRow("futai_kin")
            Cmd.Parameters(9).Value = FutaiDetailRow("futai_biko")
            Cmd.Parameters(10).Value = CommonEXT.PropComStrUserId
            Cmd.Parameters(11).Value = CommonEXT.PropComStrUserId

            Cmd.ExecuteScalar()

            '正常終了
            Return True
        Catch ex As Exception
            '例外処理
            puErrMsg = ex.Message
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 請求依頼番号の取得
    ''' </summary>
    ''' <param name="Cmd"></param>
    ''' <param name="Cn"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' <para>作成情報：2015/09/05 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function nextValSeikyuIraiNo(ByRef Cmd As NpgsqlCommand, _
                                   ByRef Cn As NpgsqlConnection) As String

        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = "SELECT 'E' || to_char(nextval('SEIKYU_IRAI_S'),'FM00000')"
        Try

            Cmd = New NpgsqlCommand(strSQL, Cn)
            Return Cmd.ExecuteScalar()
        Catch ex As Exception
            Common.CommonDeclare.puErrMsg = ex.Message
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            Return Nothing
        End Try

    End Function

    ''' <summary>
    ''' 請求入金表削除
    ''' </summary>
    ''' <param name="Cmd"></param>
    ''' <param name="Cn"></param>
    ''' <param name="YoyakuNo"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' <para>作成情報：2015/09/17 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function deleteBillPay(ByRef Cmd As NpgsqlCommand, _
                                   ByRef Cn As NpgsqlConnection, _
                                   ByVal YoyakuNo As String) As Boolean
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = String.Empty
        Try
            strSQL = strEX05D003
            Cmd = New NpgsqlCommand(strSQL, Cn)

            '値を設定
            Cmd.Parameters.Add(New NpgsqlParameter(":YoyakuNo", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters(0).Value = YoyakuNo

            Cmd.ExecuteScalar()

            '正常終了
            Return True
        Catch ex As Exception
            '例外処理
            puErrMsg = ex.Message
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            Return False
        End Try

    End Function

    ''' <summary>
    ''' 請求入金表登録
    ''' </summary>
    ''' <param name="Cmd"></param>
    ''' <param name="Cn"></param>
    ''' <param name="BillPayRow"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' <para>作成情報：2015/09/17 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function registerBillPay(ByRef Cmd As NpgsqlCommand, _
                                   ByRef Cn As NpgsqlConnection, _
                                   ByRef BillPayRow As DataRow, _
                                   ByVal YoyakuNo As String) As Boolean

        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = String.Empty
        Try
            'SQL文を設定
            BillPayRow("yoyaku_no") = YoyakuNo
            strSQL = strEX05I003
            Cmd = New NpgsqlCommand(strSQL, Cn)

            '値を設定
            Cmd.Parameters.Add(New NpgsqlParameter(":YoyakuNo", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":Seq", NpgsqlTypes.NpgsqlDbType.Integer))
            Cmd.Parameters.Add(New NpgsqlParameter(":SeikyuIraiNo", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":SeikyuDt", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":NyukinYoteiDt", NpgsqlTypes.NpgsqlDbType.Varchar))
            'Cmd.Parameters.Add(New NpgsqlParameter(":KakuteiKin", NpgsqlTypes.NpgsqlDbType.Integer))                        ' 2015.12.21 UPD h.hagiwara
            'Cmd.Parameters.Add(New NpgsqlParameter(":ChoseiKin", NpgsqlTypes.NpgsqlDbType.Integer))                         ' 2015.12.21 UPD h.hagiwara
            'Cmd.Parameters.Add(New NpgsqlParameter(":Shokei", NpgsqlTypes.NpgsqlDbType.Integer))                            ' 2015.12.21 UPD h.hagiwara
            Cmd.Parameters.Add(New NpgsqlParameter(":KakuteiKin", NpgsqlTypes.NpgsqlDbType.Bigint))                          ' 2015.12.21 UPD h.hagiwara
            Cmd.Parameters.Add(New NpgsqlParameter(":ChoseiKin", NpgsqlTypes.NpgsqlDbType.Bigint))                           ' 2015.12.21 UPD h.hagiwara
            Cmd.Parameters.Add(New NpgsqlParameter(":Shokei", NpgsqlTypes.NpgsqlDbType.Bigint))                              ' 2015.12.21 UPD h.hagiwara
            'Cmd.Parameters.Add(New NpgsqlParameter(":TaxKin", NpgsqlTypes.NpgsqlDbType.Integer))                            ' 2015.12.21 UPD h.hagiwara
            Cmd.Parameters.Add(New NpgsqlParameter(":TaxKin", NpgsqlTypes.NpgsqlDbType.Bigint))                              ' 2015.12.21 UPD h.hagiwara
            'Cmd.Parameters.Add(New NpgsqlParameter(":SeikyuKin", NpgsqlTypes.NpgsqlDbType.Integer))                         ' 2015.12.21 UPD h.hagiwara
            Cmd.Parameters.Add(New NpgsqlParameter(":SeikyuKin", NpgsqlTypes.NpgsqlDbType.Bigint))                           ' 2015.12.21 UPD h.hagiwara
            Cmd.Parameters.Add(New NpgsqlParameter(":SeikyuNaiyo", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":AiteCd", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":NyukinKbn", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":SeikyuTitle1", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":SeikyuTitle2", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":NyukinDt", NpgsqlTypes.NpgsqlDbType.Varchar))
            'Cmd.Parameters.Add(New NpgsqlParameter(":NyukinKin", NpgsqlTypes.NpgsqlDbType.Integer))                         ' 2015.12.21 UPD h.hagiwara
            Cmd.Parameters.Add(New NpgsqlParameter(":NyukinKin", NpgsqlTypes.NpgsqlDbType.Bigint))                           ' 2015.12.21 UPD h.hagiwara
            Cmd.Parameters.Add(New NpgsqlParameter(":SeikyuInputFlg", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":SeikyuIraiFlg", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":NyukinInputFlg", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":NyukinLinkNo", NpgsqlTypes.NpgsqlDbType.Integer))
            Cmd.Parameters.Add(New NpgsqlParameter(":AddUserCd", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":UpUserCd", NpgsqlTypes.NpgsqlDbType.Varchar))

            Cmd.Parameters(0).Value = YoyakuNo
            Cmd.Parameters(1).Value = BillPayRow("seq")
            Cmd.Parameters(2).Value = BillPayRow("seikyu_irai_no")
            Cmd.Parameters(3).Value = BillPayRow("seikyu_dt")
            Cmd.Parameters(4).Value = BillPayRow("nyukin_yotei_dt")
            Cmd.Parameters(5).Value = BillPayRow("kakutei_kin")
            Cmd.Parameters(6).Value = BillPayRow("chosei_kin")
            Cmd.Parameters(7).Value = BillPayRow("shokei")
            Cmd.Parameters(8).Value = BillPayRow("tax_kin")
            Cmd.Parameters(9).Value = BillPayRow("seikyu_kin")
            Cmd.Parameters(10).Value = BillPayRow("seikyu_naiyo")
            Cmd.Parameters(11).Value = BillPayRow("aite_cd")
            Cmd.Parameters(12).Value = BillPayRow("nyukin_kbn")
            Cmd.Parameters(13).Value = BillPayRow("seikyu_title1")
            Cmd.Parameters(14).Value = BillPayRow("seikyu_title2")
            Cmd.Parameters(15).Value = BillPayRow("nyukin_dt")
            Cmd.Parameters(16).Value = BillPayRow("nyukin_kin")
            Cmd.Parameters(17).Value = BillPayRow("seikyu_input_flg")
            Cmd.Parameters(18).Value = BillPayRow("seikyu_irai_flg")
            Cmd.Parameters(19).Value = BillPayRow("nyukin_input_flg")
            Cmd.Parameters(20).Value = BillPayRow("nyukin_link_no")
            Cmd.Parameters(21).Value = CommonEXT.PropComStrUserId
            Cmd.Parameters(22).Value = CommonEXT.PropComStrUserId

            Dim intSeqNo As Integer = Cmd.ExecuteScalar()
            BillPayRow("seq") = intSeqNo

            '正常終了
            Return True
        Catch ex As Exception
            '例外処理
            puErrMsg = ex.Message
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            Return False
        End Try
    End Function


    ''' <summary>
    ''' Exasプロジェクト削除(予約番号,notinSEQ)
    ''' </summary>
    ''' <param name="Cmd"></param>
    ''' <param name="Cn"></param>
    ''' <param name="YoyakuNo"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' <para>作成情報：2015/09/17 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function deleteExasProject(ByRef Cmd As NpgsqlCommand, _
                                   ByRef Cn As NpgsqlConnection, _
                                   ByVal YoyakuNo As String, _
                                   ByVal SeikyuIraiNo As String, _
                                   ByVal PrjDt As DataTable) As Boolean
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = String.Empty
        Dim strWhere As String = "("
        Try
            strSQL = strEX05D006
            ' 2015.12.25 UPD START↓ h.hagiwara
            'Dim isFirst As Boolean = True
            'Dim i As Integer = 0
            'Dim row As DataRow
            'Do While i < PrjDt.Rows.Count
            '    row = PrjDt.Rows(i)
            '    If IsDBNull(row("seq")) = False Then
            '        If isFirst = False Then
            '            strWhere = strWhere + ","
            '        Else
            '            isFirst = False
            '        End If
            '        strWhere = strWhere + row("seq").ToString
            '    End If
            '    i = i + 1
            'Loop
            'If isFirst = True Then
            '    Return True
            'End If
            'strWhere = strWhere + ")"
            'Cmd = New NpgsqlCommand(strSQL + strWhere, Cn)
            Cmd = New NpgsqlCommand(strSQL, Cn)
            ' 2015.12.25 UPD END↑ h.hagiwara

            '値を設定
            Cmd.Parameters.Add(New NpgsqlParameter(":YoyakuNo", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":SeikyuIraiNo", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters(0).Value = YoyakuNo
            Cmd.Parameters(1).Value = SeikyuIraiNo

            Cmd.ExecuteScalar()

            '正常終了
            Return True
        Catch ex As Exception
            '例外処理
            puErrMsg = ex.Message
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Cmd)
            Return False
        End Try

    End Function

    ''' <summary>
    ''' Exasプロジェクト登録・更新
    ''' </summary>
    ''' <param name="Cmd"></param>
    ''' <param name="Cn"></param>
    ''' <param name="PrjRow"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' <para>作成情報：2015/09/17 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function registerExasProject(ByRef Cmd As NpgsqlCommand, _
                                   ByRef Cn As NpgsqlConnection, _
                                   ByVal YoyakuNo As String, _
                                   ByRef PrjRow As DataRow) As Boolean

        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = String.Empty
        Try
            ' 2016.02.03 UPD START↓ h.hagiwara
            'SQL文を設定
            'If IsDBNull(PrjRow("seq")) Then
            '    'insert
            '    strSQL = strEX05I006
            'Else
            '    'update
            '    strSQL = strEX05U006
            'End If
            'insert
            strSQL = strEX05I006
            ' 2016.02.03 UPD END↑ h.hagiwara
            Cmd = New NpgsqlCommand(strSQL, Cn)

            '値を設定
            Cmd.Parameters.Add(New NpgsqlParameter(":YoyakuNo", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":Seq", NpgsqlTypes.NpgsqlDbType.Integer))
            Cmd.Parameters.Add(New NpgsqlParameter(":SeikyuIraiNo", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":SeikyuIraiSeq", NpgsqlTypes.NpgsqlDbType.Integer))
            Cmd.Parameters.Add(New NpgsqlParameter(":SeikyuNaiyo", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":RiyoYm", NpgsqlTypes.NpgsqlDbType.Varchar))
            'Cmd.Parameters.Add(New NpgsqlParameter(":KeijoKin", NpgsqlTypes.NpgsqlDbType.Integer))                                 ' 2015.12.21 UPD h.hagiwara
            Cmd.Parameters.Add(New NpgsqlParameter(":KeijoKin", NpgsqlTypes.NpgsqlDbType.Bigint))                                   ' 2015.12.21 UPD h.hagiwara
            'Cmd.Parameters.Add(New NpgsqlParameter(":TaxKin", NpgsqlTypes.NpgsqlDbType.Integer))                                   ' 2015.12.21 UPD h.hagiwara
            Cmd.Parameters.Add(New NpgsqlParameter(":TaxKin", NpgsqlTypes.NpgsqlDbType.Bigint))                                     ' 2015.12.21 UPD h.hagiwara
            Cmd.Parameters.Add(New NpgsqlParameter(":Tekiyo1", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":Tekiyo2", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":ContentCd", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":ContentUchiCd", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":ShukeiGrp", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":KamokuCd", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":SaimokuCd", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":UchiCd", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":ShosaiCd", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":KarikamokuCd", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":KariSaimokuCd", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":KariUchiCd", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":KariShosaiCd", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":AddUserCd", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":UpUserCd", NpgsqlTypes.NpgsqlDbType.Varchar))

            Cmd.Parameters(0).Value = YoyakuNo
            Cmd.Parameters(1).Value = PrjRow("seq")
            Cmd.Parameters(2).Value = PrjRow("seikyu_irai_no")
            Cmd.Parameters(3).Value = PrjRow("seikyu_irai_seq")
            Cmd.Parameters(4).Value = PrjRow("seikyu_naiyo")
            Cmd.Parameters(5).Value = PrjRow("riyo_ym")
            Cmd.Parameters(6).Value = PrjRow("keijo_kin")
            Cmd.Parameters(7).Value = PrjRow("tax_kin")
            Cmd.Parameters(8).Value = PrjRow("tekiyo1")
            Cmd.Parameters(9).Value = PrjRow("tekiyo2")
            Cmd.Parameters(10).Value = PrjRow("content_cd")
            Cmd.Parameters(11).Value = PrjRow("content_uchi_cd")
            Cmd.Parameters(12).Value = PrjRow("shukei_grp")
            Cmd.Parameters(13).Value = PrjRow("kamoku_cd")
            Cmd.Parameters(14).Value = PrjRow("saimoku_cd")
            Cmd.Parameters(15).Value = PrjRow("uchi_cd")
            Cmd.Parameters(16).Value = PrjRow("shosai_cd")
            Cmd.Parameters(17).Value = PrjRow("karikamoku_cd")
            Cmd.Parameters(18).Value = PrjRow("kari_saimoku_cd")
            Cmd.Parameters(19).Value = PrjRow("kari_uchi_cd")
            Cmd.Parameters(20).Value = PrjRow("kari_shosai_cd")
            Cmd.Parameters(21).Value = CommonEXT.PropComStrUserId
            Cmd.Parameters(22).Value = CommonEXT.PropComStrUserId

            Cmd.ExecuteScalar()
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, Nothing, Nothing, Cmd)
            '正常終了
            Return True
        Catch ex As Exception
            '例外処理
            puErrMsg = ex.Message
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 依頼履歴登録
    ''' </summary>
    ''' <param name="Cmd"></param>
    ''' <param name="Cn"></param>
    ''' <param name="row"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' <para>作成情報：2015/09/17 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function registerIraiRireki(ByRef Cmd As NpgsqlCommand, _
                                   ByRef Cn As NpgsqlConnection, _
                                   ByRef row As DataRow) As Boolean

        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = String.Empty
        Try
            strSQL = strEX05I010
            Cmd = New NpgsqlCommand(strSQL, Cn)

            '値を設定
            Cmd.Parameters.Add(New NpgsqlParameter(":YoyakuNo", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":Com", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":IraiDt", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":IraiUserCd", NpgsqlTypes.NpgsqlDbType.Varchar))

            Cmd.Parameters(0).Value = row("yoyaku_no")
            Cmd.Parameters(1).Value = row("com")
            Cmd.Parameters(2).Value = row("irai_dt")
            Cmd.Parameters(3).Value = CommonEXT.PropComStrUserId
            row("seq") = Cmd.ExecuteScalar()

            '正常終了
            Return True
        Catch ex As Exception
            '例外処理
            puErrMsg = ex.Message
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' チェック履歴登録
    ''' </summary>
    ''' <param name="Cmd"></param>
    ''' <param name="Cn"></param>
    ''' <param name="row"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' <para>作成情報：2015/09/17 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function registerCheckRireki(ByRef Cmd As NpgsqlCommand, _
                                   ByRef Cn As NpgsqlConnection, _
                                   ByRef row As DataRow) As Boolean

        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = String.Empty
        Try
            strSQL = strEX05I011
            Cmd = New NpgsqlCommand(strSQL, Cn)

            '値を設定
            Cmd.Parameters.Add(New NpgsqlParameter(":YoyakuNo", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":CheckSts", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":Com", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":CheckDt", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":CheckUserCd", NpgsqlTypes.NpgsqlDbType.Varchar))

            Cmd.Parameters(0).Value = row("yoyaku_no")
            Cmd.Parameters(1).Value = row("check_sts")
            Cmd.Parameters(2).Value = row("com")
            Cmd.Parameters(3).Value = row("check_dt")
            Cmd.Parameters(4).Value = CommonEXT.PropComStrUserId

            row("seq") = Cmd.ExecuteScalar()

            '正常終了
            Return True
        Catch ex As Exception
            '例外処理
            puErrMsg = ex.Message
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 予約情報テーブルの登録
    ''' </summary>
    ''' <param name="Cmd"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataEXTC0102"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' <para>作成情報：2015/08/05 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function registerYoyakuInfo(ByRef Cmd As NpgsqlCommand, _
                                   ByRef Cn As NpgsqlConnection, _
                                   ByVal dataEXTC0102 As DataEXTC0102) As Boolean

        Dim strSQL As String = String.Empty
        Try
            strSQL = strEX04U001
            Cmd = New NpgsqlCommand(strSQL, Cn)

            '値を設定
            Cmd.Parameters.Add(New NpgsqlParameter(":KakuteiDt", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":KakuUsercd", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":YoyakuSts", NpgsqlTypes.NpgsqlDbType.Char))
            Cmd.Parameters.Add(New NpgsqlParameter(":ShisetuKbn", NpgsqlTypes.NpgsqlDbType.Char))
            Cmd.Parameters.Add(New NpgsqlParameter(":StudioKbn", NpgsqlTypes.NpgsqlDbType.Char))
            Cmd.Parameters.Add(New NpgsqlParameter(":SaijiNm", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":ShutsuenNm", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":KashiKind", NpgsqlTypes.NpgsqlDbType.Char))
            Cmd.Parameters.Add(New NpgsqlParameter(":RiyoType", NpgsqlTypes.NpgsqlDbType.Char))
            Cmd.Parameters.Add(New NpgsqlParameter(":DrinkFlg", NpgsqlTypes.NpgsqlDbType.Char))
            Cmd.Parameters.Add(New NpgsqlParameter(":SaijiBunrui", NpgsqlTypes.NpgsqlDbType.Char))
            Cmd.Parameters.Add(New NpgsqlParameter(":Teiin", NpgsqlTypes.NpgsqlDbType.Integer))
            Cmd.Parameters.Add(New NpgsqlParameter(":OnkyoOpeFlg", NpgsqlTypes.NpgsqlDbType.Char))
            Cmd.Parameters.Add(New NpgsqlParameter(":RiyoshaCd", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":RiyoNm", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":RiyoKana", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":SekininBushoNm", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":SekininNm", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":SekininMail", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":DaihyoNm", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":RiyoTel11", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":RiyoTel12", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":RiyoTel13", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":RiyoTel21", NpgsqlTypes.NpgsqlDbType.Varchar))
            'Cmd.Parameters.Add(New NpgsqlParameter(":RiyoTel22", NpgsqlTypes.NpgsqlDbType.Varchar)) '2016.11.4 m.hayabuchi DEL 課題No.59
            'Cmd.Parameters.Add(New NpgsqlParameter(":RiyoTel23", NpgsqlTypes.NpgsqlDbType.Varchar)) '2016.11.4 m.hayabuchi DEL 課題No.59
            Cmd.Parameters.Add(New NpgsqlParameter(":RiyoNaisen", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":RiyoFax11", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":RiyoFax12", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":RiyoFax13", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":RiyoYubin1", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":RiyoYubin2", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":RiyoTodo", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":RiyoShiku", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":RiyoBan", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":RiyoBuild", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":RiyoLvl", NpgsqlTypes.NpgsqlDbType.Char))
            Cmd.Parameters.Add(New NpgsqlParameter(":AiteCd", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":OnkyoNm", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":OnkyoTantoNm", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":OnkyoTel11", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":OnkyoTel12", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":OnkyoTel13", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":OnkyoNaisen", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":OnkyoFax11", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":OnkyoFax12", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":OnkyoFax13", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":OnkyoMail", NpgsqlTypes.NpgsqlDbType.Varchar))
            'Cmd.Parameters.Add(New NpgsqlParameter(":TotalRiyoKin", NpgsqlTypes.NpgsqlDbType.Integer))                          ' 2015.12.21 UPD h.hagiwara
            Cmd.Parameters.Add(New NpgsqlParameter(":TotalRiyoKin", NpgsqlTypes.NpgsqlDbType.Bigint))                            ' 2015.12.21 UPD h.hagiwara
            Cmd.Parameters.Add(New NpgsqlParameter(":Biko", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":RiyoCom", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":TicketEnterKbn", NpgsqlTypes.NpgsqlDbType.Char))
            Cmd.Parameters.Add(New NpgsqlParameter(":TicketDrinkKbn", NpgsqlTypes.NpgsqlDbType.Char))
            Cmd.Parameters.Add(New NpgsqlParameter(":HpKeisai", NpgsqlTypes.NpgsqlDbType.Char))
            Cmd.Parameters.Add(New NpgsqlParameter(":JohoKokaiDt", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":JohoKokaiTime", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":KokaiDt", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":KokaiTime", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":FinputSts", NpgsqlTypes.NpgsqlDbType.Char))
            Cmd.Parameters.Add(New NpgsqlParameter(":AddUserCd", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":UpUserCd", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":YoyakuNo", NpgsqlTypes.NpgsqlDbType.Varchar))

            '2016.11.4 m.hayabuchi MOD Start 課題No.59
            'Cmd.Parameters(0).Value = dataEXTC0102.PropStrKakuteiDt
            'Cmd.Parameters(1).Value = dataEXTC0102.PropStrKakuUsercd
            'Cmd.Parameters(2).Value = dataEXTC0102.PropStrYoyakuSts
            'Cmd.Parameters(3).Value = dataEXTC0102.PropStrShisetuKbn
            'Cmd.Parameters(4).Value = dataEXTC0102.PropStrStudioKbn
            'Cmd.Parameters(5).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrSaijiNm)
            'Cmd.Parameters(6).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrShutsuenNm)
            'Cmd.Parameters(7).Value = dataEXTC0102.PropStrKashiKind
            'Cmd.Parameters(8).Value = DBNull.Value
            'Cmd.Parameters(9).Value = DBNull.Value
            'Cmd.Parameters(10).Value = DBNull.Value
            'Cmd.Parameters(11).Value = DBNull.Value
            'Cmd.Parameters(12).Value = dataEXTC0102.PropStrOnkyoOpe
            'Cmd.Parameters(13).Value = dataEXTC0102.PropStrRiyoshaCd
            'Cmd.Parameters(14).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrRiyoNm)
            'Cmd.Parameters(15).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrRiyoKana)
            'Cmd.Parameters(16).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrSekininBushoNm)
            'Cmd.Parameters(17).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrSekininNm)
            'Cmd.Parameters(18).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrSekininMail)
            'Cmd.Parameters(19).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrDaihyoNm)
            'Cmd.Parameters(20).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrRiyoTel11)
            'Cmd.Parameters(21).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrRiyoTel12)
            'Cmd.Parameters(22).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrRiyoTel13)
            'Cmd.Parameters(23).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrRiyoTel21)
            'Cmd.Parameters(24).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrRiyoTel22)
            'Cmd.Parameters(25).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrRiyoTel23)
            'Cmd.Parameters(26).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrRiyoNaisen)
            'Cmd.Parameters(27).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrRiyoFax11)
            'Cmd.Parameters(28).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrRiyoFax12)
            'Cmd.Parameters(29).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrRiyoFax13)
            'Cmd.Parameters(30).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrRiyoYubin1)
            'Cmd.Parameters(31).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrRiyoYubin2)
            'Cmd.Parameters(32).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrRiyoTodo)
            'Cmd.Parameters(33).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrRiyoShiku)
            'Cmd.Parameters(34).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrRiyoBan)
            'Cmd.Parameters(35).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrRiyoBuild)
            'Cmd.Parameters(36).Value = dataEXTC0102.PropStrRiyoLvl
            'Cmd.Parameters(37).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrAiteCd)
            'Cmd.Parameters(38).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrOnkyoNm)
            'Cmd.Parameters(39).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrOnkyoTantoNm)
            'Cmd.Parameters(40).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrOnkyoTel11)
            'Cmd.Parameters(41).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrOnkyoTel12)
            'Cmd.Parameters(42).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrOnkyoTel13)
            'Cmd.Parameters(43).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrOnkyoNaisen)
            'Cmd.Parameters(44).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrOnkyoFax11)
            'Cmd.Parameters(45).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrOnkyoFax12)
            'Cmd.Parameters(46).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrOnkyoFax13)
            'Cmd.Parameters(47).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrOnkyoMail)
            'Cmd.Parameters(48).Value = dataEXTC0102.propIntTtl
            'Cmd.Parameters(49).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrBiko)
            'Cmd.Parameters(50).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrRiyoCom)
            'Cmd.Parameters(51).Value = DBNull.Value
            'Cmd.Parameters(52).Value = DBNull.Value
            'Cmd.Parameters(53).Value = DBNull.Value
            'Cmd.Parameters(54).Value = DBNull.Value
            'Cmd.Parameters(55).Value = DBNull.Value
            'Cmd.Parameters(56).Value = DBNull.Value
            'Cmd.Parameters(57).Value = DBNull.Value
            'Cmd.Parameters(58).Value = dataEXTC0102.PropStrFinputSts
            'Cmd.Parameters(59).Value = CommonEXT.PropComStrUserId
            'Cmd.Parameters(60).Value = CommonEXT.PropComStrUserId
            'Cmd.Parameters(61).Value = dataEXTC0102.PropStrYoyakuNo
            Cmd.Parameters(0).Value = dataEXTC0102.PropStrKakuteiDt
            Cmd.Parameters(1).Value = dataEXTC0102.PropStrKakuUsercd
            Cmd.Parameters(2).Value = dataEXTC0102.PropStrYoyakuSts
            Cmd.Parameters(3).Value = dataEXTC0102.PropStrShisetuKbn
            Cmd.Parameters(4).Value = dataEXTC0102.PropStrStudioKbn
            Cmd.Parameters(5).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrSaijiNm)
            Cmd.Parameters(6).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrShutsuenNm)
            Cmd.Parameters(7).Value = dataEXTC0102.PropStrKashiKind
            Cmd.Parameters(8).Value = DBNull.Value
            Cmd.Parameters(9).Value = DBNull.Value
            Cmd.Parameters(10).Value = DBNull.Value
            Cmd.Parameters(11).Value = DBNull.Value
            Cmd.Parameters(12).Value = dataEXTC0102.PropStrOnkyoOpe
            Cmd.Parameters(13).Value = dataEXTC0102.PropStrRiyoshaCd
            Cmd.Parameters(14).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrRiyoNm)
            Cmd.Parameters(15).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrRiyoKana)
            Cmd.Parameters(16).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrSekininBushoNm)
            Cmd.Parameters(17).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrSekininNm)
            Cmd.Parameters(18).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrSekininMail)
            Cmd.Parameters(19).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrDaihyoNm)
            Cmd.Parameters(20).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrRiyoTel11)
            Cmd.Parameters(21).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrRiyoTel12)
            Cmd.Parameters(22).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrRiyoTel13)
            Cmd.Parameters(23).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrRiyoTel21)
            Cmd.Parameters(24).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrRiyoNaisen)
            Cmd.Parameters(25).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrRiyoFax11)
            Cmd.Parameters(26).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrRiyoFax12)
            Cmd.Parameters(27).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrRiyoFax13)
            Cmd.Parameters(28).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrRiyoYubin1)
            Cmd.Parameters(29).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrRiyoYubin2)
            Cmd.Parameters(30).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrRiyoTodo)
            Cmd.Parameters(31).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrRiyoShiku)
            Cmd.Parameters(32).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrRiyoBan)
            Cmd.Parameters(33).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrRiyoBuild)
            Cmd.Parameters(34).Value = dataEXTC0102.PropStrRiyoLvl
            Cmd.Parameters(35).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrAiteCd)
            Cmd.Parameters(36).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrOnkyoNm)
            Cmd.Parameters(37).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrOnkyoTantoNm)
            Cmd.Parameters(38).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrOnkyoTel11)
            Cmd.Parameters(39).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrOnkyoTel12)
            Cmd.Parameters(40).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrOnkyoTel13)
            Cmd.Parameters(41).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrOnkyoNaisen)
            Cmd.Parameters(42).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrOnkyoFax11)
            Cmd.Parameters(43).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrOnkyoFax12)
            Cmd.Parameters(44).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrOnkyoFax13)
            Cmd.Parameters(45).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrOnkyoMail)
            Cmd.Parameters(46).Value = dataEXTC0102.propIntTtl
            Cmd.Parameters(47).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrBiko)
            Cmd.Parameters(48).Value = commonLogicEXT.convInsStr(dataEXTC0102.PropStrRiyoCom)
            Cmd.Parameters(49).Value = DBNull.Value
            Cmd.Parameters(50).Value = DBNull.Value
            Cmd.Parameters(51).Value = DBNull.Value
            Cmd.Parameters(52).Value = DBNull.Value
            Cmd.Parameters(53).Value = DBNull.Value
            Cmd.Parameters(54).Value = DBNull.Value
            Cmd.Parameters(55).Value = DBNull.Value
            Cmd.Parameters(56).Value = dataEXTC0102.PropStrFinputSts
            Cmd.Parameters(57).Value = CommonEXT.PropComStrUserId
            Cmd.Parameters(58).Value = CommonEXT.PropComStrUserId
            Cmd.Parameters(59).Value = dataEXTC0102.PropStrYoyakuNo
            '2016.11.4 m.hayabuchi MOD End 課題No.59

            '正常終了
            Cmd.ExecuteScalar()
            Return True
        Catch ex As Exception
            '例外処理
            Common.CommonDeclare.puErrMsg = ex.Message
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Cmd)
            Return False
        End Try

    End Function

    ''' <summary>
    ''' 請求取消
    ''' </summary>
    ''' <param name="Cmd"></param>
    ''' <param name="Cn"></param>
    ''' <param name="row"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' <para>作成情報：2015/09/17 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function updateCancellSeikyu(ByRef Cmd As NpgsqlCommand, _
                                   ByRef Cn As NpgsqlConnection, _
                                   ByRef row As DataRow) As Boolean

        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = String.Empty
        Try
            'SQL文を設定
            strSQL = strEX05U001_S
            Cmd = New NpgsqlCommand(strSQL, Cn)
            '値を設定
            Cmd.Parameters.Add(New NpgsqlParameter(":YoyakuNo", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":Seq", NpgsqlTypes.NpgsqlDbType.Integer))
            Cmd.Parameters.Add(New NpgsqlParameter(":UpUserCd", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters(0).Value = row("yoyaku_no")
            Cmd.Parameters(1).Value = row("seq")
            Cmd.Parameters(2).Value = CommonEXT.PropComStrUserId
            '正常終了
            Return True
        Catch ex As Exception
            '例外処理
            puErrMsg = ex.Message
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 入金取消
    ''' </summary>
    ''' <param name="Cmd"></param>
    ''' <param name="Cn"></param>
    ''' <param name="row"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' <para>作成情報：2015/09/17 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function updateCancellNyukin(ByRef Cmd As NpgsqlCommand, _
                                   ByRef Cn As NpgsqlConnection, _
                                   ByRef row As DataRow) As Boolean

        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = String.Empty
        Try
            'SQL文を設定
            strSQL = strEX05U001_N
            Cmd = New NpgsqlCommand(strSQL, Cn)
            '値を設定
            Cmd.Parameters.Add(New NpgsqlParameter(":YoyakuNo", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":Seq", NpgsqlTypes.NpgsqlDbType.Integer))
            Cmd.Parameters.Add(New NpgsqlParameter(":UpUserCd", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters(0).Value = row("yoyaku_no")
            Cmd.Parameters(1).Value = row("seq")
            Cmd.Parameters(2).Value = CommonEXT.PropComStrUserId
            '正常終了
            Return True
        Catch ex As Exception
            '例外処理
            puErrMsg = ex.Message
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 付帯設備利用明細（合算）取得処理
    ''' </summary>
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataEXTC0103">[IN]正式予約登録/詳細画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>付帯設備利用明細（合算）取得用SQLを作成し、アダプタにセットする
    ''' <para>作成情報：2015/08/14 d.sonoda
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function SetSelectCertificateSql(ByRef Adapter As NpgsqlDataAdapter, _
                                               ByVal Cn As NpgsqlConnection, _
                                               ByVal dataEXTC0103 As DataEXTC0103) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim strSQL As String = ""

        Try
            'SQL文(SELECT)
            strSQL = strSelectCertificateSql

            'データアダプタに、SQLのSELECT文を設定
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            'バインド変数に型と値を設定
            '予約番号
            Adapter.SelectCommand.Parameters.Add("yoyakuNo", NpgsqlTypes.NpgsqlDbType.Varchar)
            Adapter.SelectCommand.Parameters("yoyakuNo").Value = dataEXTC0103.PropStrReserveNo

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, strSelectDetailsAllSql, Nothing, Nothing)

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
            Return False
        End Try

    End Function

    ''' <summary>
    ''' 付帯設備利用明細（合算）取得処理
    ''' </summary>
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataEXTC0103">[IN]正式予約登録/詳細画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>付帯設備利用明細（合算）取得用SQLを作成し、アダプタにセットする
    ''' <para>作成情報：2015/08/14 d.sonoda
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function SetSelectUseDetailsAllSql(ByRef Adapter As NpgsqlDataAdapter, _
                                               ByVal Cn As NpgsqlConnection, _
                                               ByVal dataEXTC0103 As DataEXTC0103) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim strSQL As String = ""

        Try
            'SQL文(SELECT)
            strSQL = strSelectDetailsAllSql

            'データアダプタに、SQLのSELECT文を設定
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            'バインド変数に型と値を設定
            '予約番号
            Adapter.SelectCommand.Parameters.Add("ReserveNo", NpgsqlTypes.NpgsqlDbType.Varchar)
            Adapter.SelectCommand.Parameters("ReserveNo").Value = dataEXTC0103.PropStrReserveNo
            '税無フラグ
            Adapter.SelectCommand.Parameters.Add("NoTaxFlg", NpgsqlTypes.NpgsqlDbType.Varchar)
            If dataEXTC0103.PropBlnNoTaxFlg Then
                '税無の場合
                Adapter.SelectCommand.Parameters("NoTaxFlg").Value = "1"
            Else
                '税有の場合
                Adapter.SelectCommand.Parameters("NoTaxFlg").Value = "0"
            End If

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
            Return False
        End Try

    End Function

    ''' <summary>
    ''' 付帯設備利用明細（日別）取得処理
    ''' </summary>
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataEXTC0103">[IN]正式予約登録/詳細画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>付帯設備利用明細（日別）取得用SQLを作成し、アダプタにセットする
    ''' <para>作成情報：2015/08/14 d.sonoda
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function SetSelectUseDetailsOneDaySql(ByRef Adapter As NpgsqlDataAdapter, _
                                                 ByVal Cn As NpgsqlConnection, _
                                                 ByVal dataEXTC0103 As DataEXTC0103) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim strSQL As String = ""

        Try
            'SQL文(SELECT)
            strSQL = strSelectDetailsOneDaySql

            'データアダプタに、SQLのSELECT文を設定
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            'バインド変数に型を設定
            '予約番号
            Adapter.SelectCommand.Parameters.Add("ReserveNo", NpgsqlTypes.NpgsqlDbType.Varchar)
            '税無フラグ
            Adapter.SelectCommand.Parameters.Add("NoTaxFlg", NpgsqlTypes.NpgsqlDbType.Varchar)
            '利用日
            Adapter.SelectCommand.Parameters.Add("ReserveDay", NpgsqlTypes.NpgsqlDbType.Varchar)

            'バインド変数に値を設定
            '予約番号
            Adapter.SelectCommand.Parameters("ReserveNo").Value = dataEXTC0103.PropStrReserveNo
            '税無フラグ
            If dataEXTC0103.PropBlnNoTaxFlg Then
                '税無の場合
                Adapter.SelectCommand.Parameters("NoTaxFlg").Value = "1"
            Else
                '税有の場合
                Adapter.SelectCommand.Parameters("NoTaxFlg").Value = "0"
            End If
            '利用日
            Adapter.SelectCommand.Parameters("ReserveDay").Value = dataEXTC0103.PropStrClickedIncident

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
            Return False
        End Try

    End Function




    ''' <summary>
    ''' スタジオ利用確認書SQL作成
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataHBKA0101">[IN]データクラス</param>
    ''' </summary> 
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>各テーブルから、スタジオ利用確認書のデータを取得する。
    ''' <para>作成情報：2015/08/11 yu.satoh 
    ''' </para></remarks>
    Public Function SetSelectStudioData(ByRef Adapter As NpgsqlDataAdapter, ByVal Cn As NpgsqlConnection, ByVal dataEXTC0103 As DataEXTC0103) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数の宣言
        Dim strSQL As String
        Dim strWhere As String

        Try

            'SQL文(SELECT)
            strSQL = setSelectStudio
            'Where句作成
            strWhere = " WHERE t1.yoyaku_no = :yoyakuNo"

            'データアダプタに、SQLを設定
            strSQL &= strWhere
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            'バインド変数をセット
            Adapter.SelectCommand.Parameters.Add(New NpgsqlParameter("yoyakuNo", NpgsqlTypes.NpgsqlDbType.Varchar)) '予約番号
            Adapter.SelectCommand.Parameters("yoyakuNo").Value = dataEXTC0103.PropLblYoyakuNo.Text                  '予約番号

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
            'puErrMsg = EXTM0102_E0000 & ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' 祝祭日マスタを取得するSQLの作成・設定処理
    ''' </summary>
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="CehckDay">[IN]取得対象日付</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>祝祭日マスタ取得用のSQLを作成し、アダプタにセットする
    ''' <para>作成情報：2015.10.24 h.hagiwara
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function GetHolidayMst(ByRef Adapter As NpgsqlDataAdapter, ByVal Cn As NpgsqlConnection, ByVal CehckDay As DateTime) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        '変数を宣言
        Dim strSQL As String = strHolidayCntSql

        Try
            'データアダプタに、SQLのINSERT文を設定
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)
            'バインド変数に型をセット
            Adapter.SelectCommand.Parameters.Add(New NpgsqlParameter("HOLIDAY_DT", NpgsqlTypes.NpgsqlDbType.Varchar))
            'バインド変数に値をセット
            Adapter.SelectCommand.Parameters("HOLIDAY_DT").Value = CehckDay.ToString(CommonDeclareEXT.FMT_DATE)

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常処理終了
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
    ''' 税抜フラグ取得処理
    ''' </summary>
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataEXTC0103">[IN]正式予約登録/詳細画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>付税抜フラグ取得用SQLを作成し、アダプタにセットする
    ''' <para>作成情報：2015/11/11 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function GeSqlNotaxflg(ByRef Adapter As NpgsqlDataAdapter, _
                                               ByVal Cn As NpgsqlConnection, _
                                               ByVal dataEXTC0103 As DataEXTC0103) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim strSQL As String = ""

        Try
            'SQL文(SELECT)
            strSQL = strNoTaxFlgGetSql

            'データアダプタに、SQLのSELECT文を設定
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            'バインド変数に型と値を設定
            Adapter.SelectCommand.Parameters.Add("SHUKEI_GRP", NpgsqlTypes.NpgsqlDbType.Varchar)
            Adapter.SelectCommand.Parameters.Add("KIKAN_DT", NpgsqlTypes.NpgsqlDbType.Varchar)

            Adapter.SelectCommand.Parameters("SHUKEI_GRP").Value = dataEXTC0103.PropStrGrpKey
            Adapter.SelectCommand.Parameters("KIKAN_DT").Value = dataEXTC0103.PropStrCalculateDay_Output

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
            Return False
        End Try

    End Function

    ''' <summary>
    ''' 相手先情報取得処理
    ''' </summary>
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataEXTC0103">[IN]正式予約登録/詳細画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>付税抜フラグ取得用SQLを作成し、アダプタにセットする
    ''' <para>作成情報：2015/11/11 h.hagiwara
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function GeSqlAitesaki(ByRef Adapter As NpgsqlDataAdapter, _
                                               ByVal Cn As NpgsqlConnection, _
                                               ByVal dataEXTC0103 As DataEXTC0103) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim strSQL As String = ""

        Try
            'SQL文(SELECT)
            strSQL = strAitesakiGetSql

            'データアダプタに、SQLのSELECT文を設定
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            'バインド変数に型と値を設定
            Adapter.SelectCommand.Parameters.Add("AITE_CD", NpgsqlTypes.NpgsqlDbType.Varchar)

            Adapter.SelectCommand.Parameters("AITE_CD").Value = dataEXTC0103.PropStrAitecd

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
            Return False
        End Try

    End Function

    ''' <summary>
    ''' 料金マスタを取得するSQLの作成・設定処理
    ''' </summary>
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataEXTC0103">[IN]取得対象日付</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>料金マスタ取得用のSQLを作成し、アダプタにセットする
    ''' <para>作成情報：2015.12.16 h.hagiwara
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function GeSqlRyokinMst(ByRef Adapter As NpgsqlDataAdapter, ByVal Cn As NpgsqlConnection, ByVal dataEXTC0103 As DataEXTC0103) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        '変数を宣言
        Dim strSQL As String = strRyokinGetSql

        Try
            'データアダプタに、SQLのINSERT文を設定
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)
            'バインド変数に型をセット
            Adapter.SelectCommand.Parameters.Add(New NpgsqlParameter("RIYOUBI", NpgsqlTypes.NpgsqlDbType.Varchar))
            'バインド変数に値をセット
            Adapter.SelectCommand.Parameters("RIYOUBI").Value = dataEXTC0103.PropStrClickedRiyobi

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常処理終了
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
    ''' 倍率マスタを取得するSQLの作成・設定処理
    ''' </summary>
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataEXTC0103">[IN]取得対象日付</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>倍率マスタ取得用のSQLを作成し、アダプタにセットする
    ''' <para>作成情報：2015.12.16 h.hagiwara
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function GeSqlBairituMst(ByRef Adapter As NpgsqlDataAdapter, ByVal Cn As NpgsqlConnection, ByVal dataEXTC0103 As DataEXTC0103) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        '変数を宣言
        Dim strSQL As String = strBairituGetSql

        Try
            'データアダプタに、SQLのINSERT文を設定
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)
            'バインド変数に型をセット
            Adapter.SelectCommand.Parameters.Add(New NpgsqlParameter("RIYOUBI", NpgsqlTypes.NpgsqlDbType.Varchar))
            'バインド変数に値をセット
            Adapter.SelectCommand.Parameters("RIYOUBI").Value = dataEXTC0103.PropStrClickedRiyobi

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常処理終了
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
    ''' 付帯設備マスタを取得するSQLの作成・設定処理
    ''' </summary>
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataEXTC0103">[IN]取得対象日付</param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>付帯設備マスタ取得用のSQLを作成し、アダプタにセットする
    ''' <para>作成情報：2015.12.16 h.hagiwara
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function GeSqlFutaiSetubiMst(ByRef Adapter As NpgsqlDataAdapter, ByVal Cn As NpgsqlConnection, ByVal dataEXTC0103 As DataEXTC0103) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        '変数を宣言
        Dim strSQL As String = strFutaisetubiGetSql

        Try
            'データアダプタに、SQLのINSERT文を設定
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)
            'バインド変数に型をセット
            Adapter.SelectCommand.Parameters.Add(New NpgsqlParameter("RIYOUBI", NpgsqlTypes.NpgsqlDbType.Varchar))
            'バインド変数に値をセット
            Adapter.SelectCommand.Parameters("RIYOUBI").Value = dataEXTC0103.PropStrClickedRiyobi

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            '正常処理終了
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
    ''' 責任者メールアドレス・携帯電話番号情報取得
    ''' </summary>
    ''' <param name="Adapter"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataEXTC0102"></param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>責任者名から責任者メールアドレス・携帯電話番号を取得するSQL作成
    ''' <para>作成情報：2016/08/08 e.watanabe
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function GetSqlSekininshaMailTelData(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByRef Cn As NpgsqlConnection, _
                                       ByVal dataEXTC0102 As DataEXTC0102) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = String.Empty

        Try
            'SQL文(SELECT)
            'データアダプタに、SQLのSELECT文を設定
            strSQL = strEX04S005
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            '条件項目の型を指定
            'バインド変数の型を設定
            With Adapter.SelectCommand.Parameters
                .Add(New NpgsqlParameter(":SekininNm", NpgsqlTypes.NpgsqlDbType.Varchar))
                .Add(New NpgsqlParameter(":RiyoshaCd", NpgsqlTypes.NpgsqlDbType.Varchar))
            End With
            'バインド変数の値を設定
            With Adapter.SelectCommand
                .Parameters(":SekininNm").Value = dataEXTC0102.PropStrSekininNm   '責任者名
                .Parameters(":RiyoshaCd").Value = dataEXTC0102.PropStrRiyoshaCd   '利用者コード
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

    ''' <summary>
    ''' 消費税マスタ／消費税率・軽減税率取得用SQLの作成・設定処理
    ''' </summary>
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataEXTC0103">[IN]正式予約登録/詳細画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>消費税マスタ／消費税率および軽減税率取得用SQLを作成し、アダプタにセットする
    ''' <para>作成情報：2019/08/02 E.Okuda@Compass
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function SelectTaxRateSql(ByRef Adapter As NpgsqlDataAdapter,
                                               ByVal Cn As NpgsqlConnection,
                                               ByVal dataEXTC0103 As DataEXTC0103) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim strSQL As String = ""

        Try

            'SQL文(SELECT)
            strSQL = strSelectTaxRateSql

            'データアダプタに、SQLのSELECT文を設定
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            'バインド変数に型を設定
            '利用開始日
            Adapter.SelectCommand.Parameters.Add("RiyouStart", NpgsqlTypes.NpgsqlDbType.Varchar)
            Adapter.SelectCommand.Parameters("RiyouStart").Value = dataEXTC0103.PropAryRiyouStartEnd(0)
            '利用終了日
            Adapter.SelectCommand.Parameters.Add("RiyouEnd", NpgsqlTypes.NpgsqlDbType.Varchar)
            Adapter.SelectCommand.Parameters("RiyouEnd").Value = dataEXTC0103.PropAryRiyouStartEnd(1)

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = EXT_E001 + ex.Message
            Return False
        End Try

    End Function

End Class
