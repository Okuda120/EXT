Imports Npgsql
Imports Common
Imports CommonEXT

Public Class SqlEXTY0101

    'EXAS請求依頼データ（シアター）取得SQL
    ' 2016.01.22 MOD START↓ y.morooka グループ請求対応　チェックボックス選択不可対応
    'Private strExasRequestTheaterSql As String = "SELECT FALSE AS select, " & vbCrLf & _
    '                                             "billpay_tbl.seikyu_irai_no AS seikyu_irai_no, " & vbCrLf & _
    '                                             "billpay_tbl.seikyu_dt AS seikyu_dt, " & vbCrLf & _
    '                                             "maxmin_ydt.start_day AS start_day, " & vbCrLf & _
    '                                             "maxmin_ydt.end_day AS end_day, " & vbCrLf & _
    '                                             "yoyaku_tbl.saiji_nm AS saiji_nm, " & vbCrLf & _
    '                                             "aitesaki_mst.aite_nm AS aite_nm, " & vbCrLf & _
    '                                             "billpay_tbl.seikyu_kin AS seikyu_kin, " & vbCrLf & _
    '                                             "CASE billpay_tbl.seikyu_naiyo WHEN '1' THEN '利用料' " & vbCrLf & _
    '                                             "WHEN '2' THEN '付帯設備' " & vbCrLf & _
    '                                             "WHEN '3' THEN '利用料＋付帯設備' " & vbCrLf & _
    '                                             "WHEN '4' THEN '還付' " & vbCrLf & _
    '                                             "ELSE '' END AS seikyu_naiyo, " & vbCrLf & _
    '                                             "billpay_tbl.nyukin_yotei_dt AS nyukin_yotei_dt, " & vbCrLf & _
    '                                             "yoyaku_tbl.yoyaku_no AS yoyaku_no, " & vbCrLf & _
    '                                             "billpay_tbl.seq AS seq " & vbCrLf & _
    '                                             "FROM yoyaku_tbl " & vbCrLf & _
    '                                             "LEFT OUTER JOIN billpay_tbl " & vbCrLf & _
    '                                             "ON yoyaku_tbl.yoyaku_no = billpay_tbl.yoyaku_no " & vbCrLf & _
    '                                             "LEFT OUTER JOIN aitesaki_mst " & vbCrLf & _
    '                                             "ON yoyaku_tbl.aite_cd = aitesaki_mst.aite_cd " & vbCrLf & _
    '                                             "LEFT OUTER JOIN ( " & vbCrLf & _
    '                                             "SELECT ydt_tbl.yoyaku_no, " & vbCrLf & _
    '                                             "MIN(ydt_tbl.yoyaku_dt) AS start_day, " & vbCrLf & _
    '                                             "MAX(ydt_tbl.yoyaku_dt) AS end_day " & vbCrLf & _
    '                                             "FROM ydt_tbl " & vbCrLf & _
    '                                             "GROUP BY yoyaku_no " & vbCrLf & _
    '                                             ") maxmin_ydt " & vbCrLf & _
    '                                             "ON yoyaku_tbl.yoyaku_no = maxmin_ydt.yoyaku_no  " & vbCrLf & _
    '                                             "WHERE yoyaku_tbl.shisetu_kbn = '1' " & vbCrLf & _
    '                                             "AND billpay_tbl.seikyu_input_flg = '1' " & vbCrLf & _
    '                                             "AND billpay_tbl.seikyu_irai_flg = '0' " & vbCrLf & _
    '                                             "AND billpay_tbl.seikyu_kin <> 0 " & vbCrLf & _
    '                                             "ORDER BY billpay_tbl.seikyu_irai_no "

    Private strExasRequestTheaterSql As String = "SELECT FALSE AS select, " & vbCrLf & _
                                                "billpay_tbl.seikyu_irai_no AS seikyu_irai_no, " & vbCrLf & _
                                                "billpay_tbl.seikyu_dt AS seikyu_dt, " & vbCrLf & _
                                                "maxmin_ydt.start_day AS start_day, " & vbCrLf & _
                                                "maxmin_ydt.end_day AS end_day, " & vbCrLf & _
                                                "yoyaku_tbl.saiji_nm AS saiji_nm, " & vbCrLf & _
                                                "aitesaki_mst.aite_nm AS aite_nm, " & vbCrLf & _
                                                "billpay_tbl.seikyu_kin AS seikyu_kin, " & vbCrLf & _
                                                "CASE billpay_tbl.seikyu_naiyo WHEN '1' THEN '利用料' " & vbCrLf & _
                                                "WHEN '2' THEN '付帯設備' " & vbCrLf & _
                                                "WHEN '3' THEN '利用料＋付帯設備' " & vbCrLf & _
                                                "WHEN '4' THEN '還付' " & vbCrLf & _
                                                "ELSE '' END AS seikyu_naiyo, " & vbCrLf & _
                                                "billpay_tbl.nyukin_yotei_dt AS nyukin_yotei_dt, " & vbCrLf & _
                                                "yoyaku_tbl.yoyaku_no AS yoyaku_no, " & vbCrLf & _
                                                "billpay_tbl.seq AS seq, " & vbCrLf & _
                                                "yoyaku_tbl.aite_cd AS aite_cd " & vbCrLf & _
                                                "FROM yoyaku_tbl " & vbCrLf & _
                                                "LEFT OUTER JOIN billpay_tbl " & vbCrLf & _
                                                "ON yoyaku_tbl.yoyaku_no = billpay_tbl.yoyaku_no " & vbCrLf & _
                                                "LEFT OUTER JOIN aitesaki_mst " & vbCrLf & _
                                                "ON yoyaku_tbl.aite_cd = aitesaki_mst.aite_cd " & vbCrLf & _
                                                "LEFT OUTER JOIN ( " & vbCrLf & _
                                                "SELECT ydt_tbl.yoyaku_no, " & vbCrLf & _
                                                "MIN(ydt_tbl.yoyaku_dt) AS start_day, " & vbCrLf & _
                                                "MAX(ydt_tbl.yoyaku_dt) AS end_day " & vbCrLf & _
                                                "FROM ydt_tbl " & vbCrLf & _
                                                "GROUP BY yoyaku_no " & vbCrLf & _
                                                ") maxmin_ydt " & vbCrLf & _
                                                "ON yoyaku_tbl.yoyaku_no = maxmin_ydt.yoyaku_no  " & vbCrLf & _
                                                "WHERE yoyaku_tbl.shisetu_kbn = '1' " & vbCrLf & _
                                                "AND billpay_tbl.seikyu_input_flg = '1' " & vbCrLf & _
                                                "AND billpay_tbl.seikyu_irai_flg = '0' " & vbCrLf & _
                                                "AND billpay_tbl.seikyu_kin <> 0 " & vbCrLf & _
                                                "ORDER BY billpay_tbl.seikyu_irai_no "
    '2016.01.22 MOD END ↑ y.morooka グループ請求対応　　チェックボックス選択不可対応
    'EXAS請求依頼データ（スタジオ）取得SQL
    ' 2016.01.22 MOD START↓ y.morooka グループ請求対応　チェックボックス選択不可対応
    'Private strExasRequestStudioSql As String = "SELECT FALSE AS select, " & vbCrLf & _
    '                                            "billpay_tbl.seikyu_irai_no AS seikyu_irai_no, " & vbCrLf & _
    '                                            "billpay_tbl.seikyu_dt AS seikyu_dt, " & vbCrLf & _
    '                                            "maxmin_ydt.start_day AS start_day, " & vbCrLf & _
    '                                            "maxmin_ydt.end_day AS end_day, " & vbCrLf & _
    '                                            "CASE yoyaku_tbl.studio_kbn WHEN '1' THEN '201st' " & vbCrLf & _
    '                                            "WHEN '2' THEN '202st' " & vbCrLf & _
    '                                            "WHEN '3' THEN 'house lock' " & vbCrLf & _
    '                                            "ELSE '' END AS studio_kbn, " & vbCrLf & _
    '                                            "yoyaku_tbl.shutsuen_nm AS artist_name, " & vbCrLf & _
    '                                            "aitesaki_mst.aite_nm AS aite_nm, " & vbCrLf & _
    '                                            "billpay_tbl.seikyu_kin AS seikyu_kin, " & vbCrLf & _
    '                                            "CASE billpay_tbl.seikyu_naiyo WHEN '1' THEN '利用料' " & vbCrLf & _
    '                                            "WHEN '2' THEN '付帯設備' " & vbCrLf & _
    '                                            "WHEN '3' THEN '利用料＋付帯設備' " & vbCrLf & _
    '                                            "WHEN '4' THEN '還付' " & vbCrLf & _
    '                                            "ELSE '' END AS seikyu_naiyo, " & vbCrLf & _
    '                                            "billpay_tbl.nyukin_yotei_dt AS nyukin_yotei_dt, " & vbCrLf & _
    '                                            "yoyaku_tbl.yoyaku_no AS yoyaku_no, " & vbCrLf & _
    '                                             "billpay_tbl.seq AS seq " & vbCrLf & _
    '                                            "FROM yoyaku_tbl " & vbCrLf & _
    '                                            "LEFT OUTER JOIN billpay_tbl " & vbCrLf & _
    '                                            "ON yoyaku_tbl.yoyaku_no = billpay_tbl.yoyaku_no " & vbCrLf & _
    '                                            "LEFT OUTER JOIN aitesaki_mst " & vbCrLf & _
    '                                            "ON yoyaku_tbl.aite_cd = aitesaki_mst.aite_cd " & vbCrLf & _
    '                                            "LEFT OUTER JOIN ( " & vbCrLf & _
    '                                            "SELECT ydt_tbl.yoyaku_no, " & vbCrLf & _
    '                                            "MIN(ydt_tbl.yoyaku_dt) AS start_day, " & vbCrLf & _
    '                                            "MAX(ydt_tbl.yoyaku_dt) AS end_day " & vbCrLf & _
    '                                            "FROM ydt_tbl " & vbCrLf & _
    '                                            "GROUP BY yoyaku_no " & vbCrLf & _
    '                                            ") maxmin_ydt " & vbCrLf & _
    '                                            "ON yoyaku_tbl.yoyaku_no = maxmin_ydt.yoyaku_no  " & vbCrLf & _
    '                                            "WHERE yoyaku_tbl.shisetu_kbn = '2' " & vbCrLf & _
    '                                            "AND billpay_tbl.seikyu_input_flg = '1' " & vbCrLf & _
    '                                            "AND billpay_tbl.seikyu_irai_flg = '0' " & vbCrLf & _
    '                                            "AND billpay_tbl.seikyu_kin <> 0 " & vbCrLf & _
    '                                            "ORDER BY billpay_tbl.seikyu_irai_no "

    Private strExasRequestStudioSql As String = "SELECT FALSE AS select, " & vbCrLf & _
                                               "billpay_tbl.seikyu_irai_no AS seikyu_irai_no, " & vbCrLf & _
                                               "billpay_tbl.seikyu_dt AS seikyu_dt, " & vbCrLf & _
                                               "maxmin_ydt.start_day AS start_day, " & vbCrLf & _
                                               "maxmin_ydt.end_day AS end_day, " & vbCrLf & _
                                               "CASE yoyaku_tbl.studio_kbn WHEN '1' THEN '201st' " & vbCrLf & _
                                               "WHEN '2' THEN '202st' " & vbCrLf & _
                                               "WHEN '3' THEN 'house lock' " & vbCrLf & _
                                               "ELSE '' END AS studio_kbn, " & vbCrLf & _
                                               "yoyaku_tbl.shutsuen_nm AS artist_name, " & vbCrLf & _
                                               "aitesaki_mst.aite_nm AS aite_nm, " & vbCrLf & _
                                               "billpay_tbl.seikyu_kin AS seikyu_kin, " & vbCrLf & _
                                               "CASE billpay_tbl.seikyu_naiyo WHEN '1' THEN '利用料' " & vbCrLf & _
                                               "WHEN '2' THEN '付帯設備' " & vbCrLf & _
                                               "WHEN '3' THEN '利用料＋付帯設備' " & vbCrLf & _
                                               "WHEN '4' THEN '還付' " & vbCrLf & _
                                               "ELSE '' END AS seikyu_naiyo, " & vbCrLf & _
                                               "billpay_tbl.nyukin_yotei_dt AS nyukin_yotei_dt, " & vbCrLf & _
                                               "yoyaku_tbl.yoyaku_no AS yoyaku_no, " & vbCrLf & _
                                               "billpay_tbl.seq AS seq, " & vbCrLf & _
                                               "yoyaku_tbl.aite_cd AS aite_cd " & vbCrLf & _
                                               "FROM yoyaku_tbl " & vbCrLf & _
                                               "LEFT OUTER JOIN billpay_tbl " & vbCrLf & _
                                               "ON yoyaku_tbl.yoyaku_no = billpay_tbl.yoyaku_no " & vbCrLf & _
                                               "LEFT OUTER JOIN aitesaki_mst " & vbCrLf & _
                                               "ON yoyaku_tbl.aite_cd = aitesaki_mst.aite_cd " & vbCrLf & _
                                               "LEFT OUTER JOIN ( " & vbCrLf & _
                                               "SELECT ydt_tbl.yoyaku_no, " & vbCrLf & _
                                               "MIN(ydt_tbl.yoyaku_dt) AS start_day, " & vbCrLf & _
                                               "MAX(ydt_tbl.yoyaku_dt) AS end_day " & vbCrLf & _
                                               "FROM ydt_tbl " & vbCrLf & _
                                               "GROUP BY yoyaku_no " & vbCrLf & _
                                               ") maxmin_ydt " & vbCrLf & _
                                               "ON yoyaku_tbl.yoyaku_no = maxmin_ydt.yoyaku_no  " & vbCrLf & _
                                               "WHERE yoyaku_tbl.shisetu_kbn = '2' " & vbCrLf & _
                                               "AND billpay_tbl.seikyu_input_flg = '1' " & vbCrLf & _
                                               "AND billpay_tbl.seikyu_irai_flg = '0' " & vbCrLf & _
                                               "AND billpay_tbl.seikyu_kin <> 0 " & vbCrLf & _
                                               "ORDER BY billpay_tbl.seikyu_irai_no "
    '2016.01.22 MOD END ↑ y.morooka グループ請求対応　　チェックボックス選択不可対応

    'CSV出力用EXAS請求依頼データ取得SQL
    'Private strSelectCsvDataSql As String = "SELECT t1.seikyu_dt, " & vbCrLf &
    '                                        "t1.nyukin_yotei_dt, " & vbCrLf &
    '                                        "t1.aite_cd, " & vbCrLf &
    '                                        "t2.post_bango, " & vbCrLf &
    '                                        "t2.add1, " & vbCrLf &
    '                                        "t2.add2, " & vbCrLf &
    '                                        "t2.aite_nm, " & vbCrLf &
    '                                        "t1.seikyu_title1, " & vbCrLf &
    '                                        "t1.seikyu_title2, " & vbCrLf &
    '                                        "t3.content_cd, " & vbCrLf &
    '                                        "t3.content_uchi_cd, " & vbCrLf &
    '                                        "t3.kamoku_cd, " & vbCrLf &
    '                                        "t3.saimoku_cd, " & vbCrLf &
    '                                        "t3.uchi_cd, " & vbCrLf &
    '                                        "t3.shosai_cd, " & vbCrLf &
    '                                        "t3.karikamoku_cd, " & vbCrLf &
    '                                        "t3.kari_saimoku_cd, " & vbCrLf &
    '                                        "t3.kari_uchi_cd, " & vbCrLf &
    '                                        "t3.kari_shosai_cd, " & vbCrLf &
    '                                        "t3.riyo_ym, " & vbCrLf &
    '                                        "t3.tekiyo1, " & vbCrLf &
    '                                        "t3.tekiyo2, " & vbCrLf &
    '                                        "t3.seikyu_irai_no, " & vbCrLf &
    '                                        "t3.keijo_kin, " & vbCrLf &
    '                                        "t3.tax_kin, " & vbCrLf &
    '                                        "t5.tax_ritu " & vbCrLf &
    '                                        "FROM billpay_tbl t1 " & vbCrLf &
    '                                        "LEFT JOIN aitesaki_mst t2 " & vbCrLf &
    '                                        "ON t1.aite_cd = t2.aite_cd " & vbCrLf &
    '                                        "LEFT JOIN project_tbl t3 " & vbCrLf &
    '                                        "ON t1.yoyaku_no = t3.yoyaku_no " & vbCrLf &
    '                                        "AND t1.seq = t3.seq " & vbCrLf &
    '                                        "AND t1.seikyu_irai_no = t3.seikyu_irai_no " & vbCrLf &
    '                                        "LEFT JOIN ( " & vbCrLf &
    '                                        "SELECT yoyaku_no, " & vbCrLf &
    '                                        "MIN(yoyaku_dt) AS min_dt " & vbCrLf &
    '                                        "FROM ydt_tbl " & vbCrLf &
    '                                        "GROUP BY yoyaku_no " & vbCrLf &
    '                                        ") t4 " & vbCrLf &
    '                                        "ON t1.yoyaku_no = t4.yoyaku_no " & vbCrLf &
    '                                        "LEFT JOIN tax_mst t5 " & vbCrLf &
    '                                        "ON TO_DATE(t4.min_dt, 'yyyy/MM/dd') > TO_DATE(t5.taxs_dt, 'yyyy/MM/dd') " & vbCrLf &
    '                                        "AND TO_DATE(t4.min_dt, 'yyyy/MM/dd') < TO_DATE(t5.taxe_dt, 'yyyy/MM/dd')  " & vbCrLf &
    '                                        "WHERE t1.yoyaku_no = :ReserveNo " & vbCrLf &
    '                                        "AND t1.seq = :Seq" & vbCrLf &
    '                                        "AND t1.seikyu_irai_no = :BillNo"
    ' 2016.03.17 UPD START↓ h.hagiwara 請求内容（利用料、付帯設備）の取得追加
    'Private strSelectCsvDataSql As String = "SELECT " & vbCrLf & _
    '                                        "  replace(t1.seikyu_dt,'/',''), " & vbCrLf & _
    '                                        "  replace(t1.nyukin_yotei_dt,'/',''), " & vbCrLf & _
    '                                        "  t1.aite_cd, " & vbCrLf & _
    '                                        "  t2.post_bango, " & vbCrLf & _
    '                                        "  t2.add1, " & vbCrLf & _
    '                                        "  t2.add2, " & vbCrLf & _
    '                                        "  t2.aite_nm, " & vbCrLf & _
    '                                        "  t1.seikyu_title1, " & vbCrLf & _
    '                                        "  t1.seikyu_title2, " & vbCrLf & _
    '                                        "  t3.content_cd, " & vbCrLf & _
    '                                        "  t3.content_uchi_cd, " & vbCrLf & _
    '                                        "  t3.kamoku_cd, " & vbCrLf & _
    '                                        "  t3.saimoku_cd, " & vbCrLf & _
    '                                        "  t3.uchi_cd, " & vbCrLf & _
    '                                        "  t3.shosai_cd, " & vbCrLf & _
    '                                        "  t3.karikamoku_cd, " & vbCrLf & _
    '                                        "  t3.kari_saimoku_cd, " & vbCrLf & _
    '                                        "  t3.kari_uchi_cd, " & vbCrLf & _
    '                                        "  t3.kari_shosai_cd, " & vbCrLf & _
    '                                        "  t3.riyo_ym, " & vbCrLf & _
    '                                        "  t3.tekiyo1, " & vbCrLf & _
    '                                        "  t3.tekiyo2, " & vbCrLf & _
    '                                        "  t3.seikyu_irai_no, " & vbCrLf & _
    '                                        "  t3.keijo_kin, " & vbCrLf & _
    '                                        "  t3.tax_kin, " & vbCrLf & _
    '                                        "  CASE WHEN COALESCE(t6.notax_flg,'0') = '1' THEN 0 " & vbCrLf & _
    '                                        "       ELSE t5.tax_ritu " & vbCrLf & _
    '                                        "  END tax_ritu ," & vbCrLf & _
    '                                        "  COALESCE(t6.notax_flg,'0') " & vbCrLf & _
    '                                        "FROM billpay_tbl t1 " & vbCrLf & _
    '                                        "  LEFT JOIN aitesaki_mst t2 " & vbCrLf & _
    '                                        "        ON t1.aite_cd = t2.aite_cd " & vbCrLf & _
    '                                        "  LEFT JOIN project_tbl t3 " & vbCrLf & _
    '                                        "        ON t1.yoyaku_no = t3.yoyaku_no " & vbCrLf & _
    '                                        "       AND t1.seikyu_irai_no = t3.seikyu_irai_no " & vbCrLf & _
    '                                        "  LEFT JOIN ( " & vbCrLf & _
    '                                        "              SELECT yoyaku_no, " & vbCrLf & _
    '                                        "                     shisetu_kbn , " & vbCrLf & _
    '                                        "                     MIN(yoyaku_dt) AS min_dt " & vbCrLf & _
    '                                        "              FROM ydt_tbl " & vbCrLf & _
    '                                        "              GROUP BY yoyaku_no ,shisetu_kbn" & vbCrLf & _
    '                                        "            ) t4 " & vbCrLf & _
    '                                        "        ON t1.yoyaku_no = t4.yoyaku_no " & vbCrLf & _
    '                                        "  LEFT JOIN tax_mst t5 " & vbCrLf & _
    '                                        "        ON  TO_DATE(t4.min_dt, 'yyyy/MM/dd') >= TO_DATE(t5.taxs_dt, 'yyyy/MM/dd') " & vbCrLf & _
    '                                        "       AND  TO_DATE(t4.min_dt, 'yyyy/MM/dd') <= TO_DATE(t5.taxe_dt, 'yyyy/MM/dd')  " & vbCrLf & _
    '                                        "  LEFT JOIN ( " & vbCrLf & _
    '                                        "              SELECT kikan_from , kikan_to , shisetu_kbn , shukei_grp , MAX(notax_flg) AS notax_flg " & vbCrLf & _
    '                                        "              FROM FBUNRUI_MST " & vbCrLf & _
    '                                        "              GROUP BY kikan_from , kikan_to , shisetu_kbn , shukei_grp " & vbCrLf & _
    '                                        "            ) t6 " & vbCrLf & _
    '                                        "        ON  t3.shukei_grp = t6.shukei_grp " & vbCrLf & _
    '                                        "       AND  TO_DATE(t4.min_dt, 'yyyy/MM/dd') >= TO_DATE(t6.kikan_from, 'yyyy/MM/dd')  " & vbCrLf & _
    '                                        "       AND  TO_DATE(t4.min_dt, 'yyyy/MM/dd') <= TO_DATE(t6.kikan_to, 'yyyy/MM/dd')  " & vbCrLf & _
    '                                        "       AND  t4.shisetu_kbn = t6.shisetu_kbn " & vbCrLf & _
    '                                        "WHERE t1.yoyaku_no = :ReserveNo " & vbCrLf & _
    '                                        "  AND t1.seq = :Seq" & vbCrLf & _
    '                                        "  AND t1.seikyu_irai_no = :BillNo"
    Private strSelectCsvDataSql As String = "SELECT " & vbCrLf & _
                                            "  replace(t1.seikyu_dt,'/',''), " & vbCrLf & _
                                            "  replace(t1.nyukin_yotei_dt,'/',''), " & vbCrLf & _
                                            "  t1.aite_cd, " & vbCrLf & _
                                            "  t2.post_bango, " & vbCrLf & _
                                            "  t2.add1, " & vbCrLf & _
                                            "  t2.add2, " & vbCrLf & _
                                            "  t2.aite_nm, " & vbCrLf & _
                                            "  t1.seikyu_title1, " & vbCrLf & _
                                            "  t1.seikyu_title2, " & vbCrLf & _
                                            "  t3.content_cd, " & vbCrLf & _
                                            "  t3.content_uchi_cd, " & vbCrLf & _
                                            "  t3.kamoku_cd, " & vbCrLf & _
                                            "  t3.saimoku_cd, " & vbCrLf & _
                                            "  t3.uchi_cd, " & vbCrLf & _
                                            "  t3.shosai_cd, " & vbCrLf & _
                                            "  t3.karikamoku_cd, " & vbCrLf & _
                                            "  t3.kari_saimoku_cd, " & vbCrLf & _
                                            "  t3.kari_uchi_cd, " & vbCrLf & _
                                            "  t3.kari_shosai_cd, " & vbCrLf & _
                                            "  t3.riyo_ym, " & vbCrLf & _
                                            "  t3.tekiyo1, " & vbCrLf & _
                                            "  t3.tekiyo2, " & vbCrLf & _
                                            "  t3.seikyu_irai_no, " & vbCrLf & _
                                            "  t3.keijo_kin, " & vbCrLf & _
                                            "  t3.tax_kin, " & vbCrLf & _
                                            "  CASE WHEN t3.SEIKYU_NAIYO = '1' THEN t5.tax_ritu " & vbCrLf & _
                                            "       ELSE CASE  " & vbCrLf & _
                                            "               WHEN COALESCE(t6.notax_flg,'0') = '1' THEN 0 " & vbCrLf & _
                                            "               ELSE t5.tax_ritu " & vbCrLf & _
                                            "            END " & vbCrLf & _
                                            "  END tax_ritu ," & vbCrLf & _
                                            "  COALESCE(t6.notax_flg,'0') ," & vbCrLf & _
                                            "  t3.SEIKYU_NAIYO " & vbCrLf & _
                                            "FROM billpay_tbl t1 " & vbCrLf & _
                                            "  LEFT JOIN aitesaki_mst t2 " & vbCrLf & _
                                            "        ON t1.aite_cd = t2.aite_cd " & vbCrLf & _
                                            "  LEFT JOIN project_tbl t3 " & vbCrLf & _
                                            "        ON t1.yoyaku_no = t3.yoyaku_no " & vbCrLf & _
                                            "       AND t1.seikyu_irai_no = t3.seikyu_irai_no " & vbCrLf & _
                                            "  LEFT JOIN ( " & vbCrLf & _
                                            "              SELECT yoyaku_no, " & vbCrLf & _
                                            "                     shisetu_kbn , " & vbCrLf & _
                                            "                     MIN(yoyaku_dt) AS min_dt " & vbCrLf & _
                                            "              FROM ydt_tbl " & vbCrLf & _
                                            "              GROUP BY yoyaku_no ,shisetu_kbn" & vbCrLf & _
                                            "            ) t4 " & vbCrLf & _
                                            "        ON t1.yoyaku_no = t4.yoyaku_no " & vbCrLf & _
                                            "  LEFT JOIN tax_mst t5 " & vbCrLf & _
                                            "        ON  TO_DATE(t4.min_dt, 'yyyy/MM/dd') >= TO_DATE(t5.taxs_dt, 'yyyy/MM/dd') " & vbCrLf & _
                                            "       AND  TO_DATE(t4.min_dt, 'yyyy/MM/dd') <= TO_DATE(t5.taxe_dt, 'yyyy/MM/dd')  " & vbCrLf & _
                                            "  LEFT JOIN ( " & vbCrLf & _
                                            "              SELECT kikan_from , kikan_to , shisetu_kbn , shukei_grp , MAX(notax_flg) AS notax_flg " & vbCrLf & _
                                            "              FROM FBUNRUI_MST " & vbCrLf & _
                                            "              GROUP BY kikan_from , kikan_to , shisetu_kbn , shukei_grp " & vbCrLf & _
                                            "            ) t6 " & vbCrLf & _
                                            "        ON  t3.shukei_grp = t6.shukei_grp " & vbCrLf & _
                                            "       AND  TO_DATE(t4.min_dt, 'yyyy/MM/dd') >= TO_DATE(t6.kikan_from, 'yyyy/MM/dd')  " & vbCrLf & _
                                            "       AND  TO_DATE(t4.min_dt, 'yyyy/MM/dd') <= TO_DATE(t6.kikan_to, 'yyyy/MM/dd')  " & vbCrLf & _
                                            "       AND  t4.shisetu_kbn = t6.shisetu_kbn " & vbCrLf & _
                                            "WHERE t1.yoyaku_no = :ReserveNo " & vbCrLf & _
                                            "  AND t1.seq = :Seq" & vbCrLf & _
                                            "  AND t1.seikyu_irai_no = :BillNo"
    ' 2016.03.17 UPD END↑ h.hagiwara 請求内容（利用料、付帯設備）の取得追加

    'EXAS請求依頼済フラグ更新SQL
    'Private strUpdateSeikyuIraiFlgSql As String = "UPDATE billpay_tbl" & vbCrLf &
    '                                              "SET seikyu_irai_flg = '1'," & vbCrLf &
    '                                              "up_dt = now()" & vbCrLf &
    '                                              "WHERE yoyaku_no = :ReserveNo " & vbCrLf &
    '                                              "AND seq = :Seq" & vbCrLf &
    '                                              "AND seikyu_irai_no = :BillNo"
    Private strUpdateSeikyuIraiFlgSql As String = "UPDATE billpay_tbl" & vbCrLf &
                                                  "SET seikyu_irai_flg = '1'," & vbCrLf &
                                                  "up_dt = now()" & vbCrLf &
                                                  "WHERE yoyaku_no = :ReserveNo " & vbCrLf &
                                                  "AND seikyu_irai_no = :BillNo"

    ''' <summary>
    ''' EXAS請求依頼データ（シアター）取得SQLの作成・設定処理
    ''' </summary>
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataEXTY0101">[IN]EXAS請求依頼データ作成画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>EXAS請求依頼データ（シアター）取得SQLを作成し、アダプタにセットする
    ''' <para>作成情報：2015/08/25 d.sonoda
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function SetExasRequestTheaterSql(ByRef Adapter As NpgsqlDataAdapter, _
                                             ByVal Cn As NpgsqlConnection, _
                                             ByVal dataEXTY0101 As DataEXTY0101) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim strSQL As String = ""

        Try
            'SQL文(SELECT)
            strSQL = strExasRequestTheaterSql

            'データアダプタに、SQLのSELECT文を設定
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

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

    ''' <summary>
    ''' EXAS請求依頼データ（スタジオ）取得SQLの作成・設定処理
    ''' </summary>
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataEXTY0101">[IN]EXAS請求依頼データ作成画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>EXAS請求依頼データ（スタジオ）取得SQLを作成し、アダプタにセットする
    ''' <para>作成情報：2015/08/25 d.sonoda
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function SetExasRequestStudioSql(ByRef Adapter As NpgsqlDataAdapter, _
                                             ByVal Cn As NpgsqlConnection, _
                                             ByVal dataEXTY0101 As DataEXTY0101) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim strSQL As String = ""

        Try
            'SQL文(SELECT)
            strSQL = strExasRequestStudioSql

            'データアダプタに、SQLのSELECT文を設定
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

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

    ''' <summary>
    ''' CSV出力用EXAS請求依頼データ取得SQLの作成・設定処理
    ''' </summary>
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataEXTY0101">[IN]EXAS請求依頼データ作成画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>CSV出力用EXAS請求依頼データ取得SQLを作成し、アダプタにセットする
    ''' <para>作成情報：2015/08/25 d.sonoda
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function SetOutputCsvDataSql(ByRef Adapter As NpgsqlDataAdapter, _
                                         ByVal Cn As NpgsqlConnection, _
                                         ByVal dataEXTY0101 As DataEXTY0101) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim strSQL As String = ""

        Try
            'SQL文(SELECT)
            strSQL = strSelectCsvDataSql

            'データアダプタに、SQLのSELECT文を設定
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            'バインド変数に型を設定
            Adapter.SelectCommand.Parameters.Add("ReserveNo", NpgsqlTypes.NpgsqlDbType.Varchar)
            Adapter.SelectCommand.Parameters.Add("Seq", NpgsqlTypes.NpgsqlDbType.Integer)
            Adapter.SelectCommand.Parameters.Add("BillNo", NpgsqlTypes.NpgsqlDbType.Varchar)

            'バインド変数に値を設定
            Adapter.SelectCommand.Parameters("ReserveNo").Value = dataEXTY0101.PropstrReserveNo_Search
            Adapter.SelectCommand.Parameters("Seq").Value = dataEXTY0101.PropStrSeq_Search
            Adapter.SelectCommand.Parameters("BillNo").Value = dataEXTY0101.PropStrBillNo_Search

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

    ''' <summary>
    ''' EXAS請求依頼済フラグ更新SQLの作成・設定処理
    ''' </summary>
    ''' <param name="Cmd">[IN/OUT]NpgSqlCommandクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataEXTY0101">[IN]EXAS請求依頼データ作成画面Dataクラス</param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>EXAS請求依頼済フラグ更新SQLを作成し、アダプタにセットする
    ''' <para>作成情報：2015/08/25 d.sonoda
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function SetUpdateSeikyuIraiFlgSql(ByRef Cmd As NpgsqlCommand, _
                                              ByVal Cn As NpgsqlConnection, _
                                              ByVal dataEXTY0101 As DataEXTY0101) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim strSQL As String = ""
        Dim strWhere As String = ""

        Try
            'SQL文(SELECT)
            strSQL = strUpdateSeikyuIraiFlgSql

            'データアダプタに、SQLのSELECT文を設定
            Cmd = New NpgsqlCommand(strSQL, Cn)

            'バインド変数に型を設定
            Cmd.Parameters.Add(New NpgsqlParameter("ReserveNo", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter("Seq", NpgsqlTypes.NpgsqlDbType.Integer))
            Cmd.Parameters.Add(New NpgsqlParameter("BillNo", NpgsqlTypes.NpgsqlDbType.Varchar))

            'バインド変数に値を設定
            Cmd.Parameters("ReserveNo").Value = dataEXTY0101.PropstrReserveNo_Search
            Cmd.Parameters("Seq").Value = dataEXTY0101.PropStrSeq_Search
            Cmd.Parameters("BillNo").Value = dataEXTY0101.PropStrBillNo_Search

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Cmd)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = EXT_E001 + ex.Message
            Return False
        End Try

    End Function

End Class
