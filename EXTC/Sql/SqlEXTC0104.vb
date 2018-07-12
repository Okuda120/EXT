Imports Common
Imports Npgsql
Imports CommonEXT
Imports EXTZ

Public Class SqlEXTC0104

    Private commonLogicEXT As New CommonLogicEXT

    'SQL文(キャンセル待ち情報の取得)
    '2016.11.4 m.hayabuchi MOD Start 課題No.59
    'Private strEX06S001 As String = _
    '                        "select " & vbCrLf & _
    '                           "t1.cancel_wait_no, " & vbCrLf & _
    '                           "t1.cancel_wait_dt, " & vbCrLf & _
    '                           "t1.cancel_wait_usercd, " & vbCrLf & _
    '                           "t1.cancel_wait_sts, " & vbCrLf & _
    '                           "t1.shisetu_kbn, " & vbCrLf & _
    '                           "t1.studio_kbn, " & vbCrLf & _
    '                           "t1.saiji_nm, " & vbCrLf & _
    '                           "t1.shutsuen_nm, " & vbCrLf & _
    '                           "t1.kashi_kind, " & vbCrLf & _
    '                           "t1.riyo_type, " & vbCrLf & _
    '                           "t1.drink_flg, " & vbCrLf & _
    '                           "t1.saiji_bunrui, " & vbCrLf & _
    '                           "t1.teiin, " & vbCrLf & _
    '                           "t1.onkyo_ope_flg, " & vbCrLf & _
    '                           "t1.riyosha_cd, " & vbCrLf & _
    '                           "t1.riyo_nm, " & vbCrLf & _
    '                           "t1.riyo_kana, " & vbCrLf & _
    '                           "t1.sekinin_busho_nm, " & vbCrLf & _
    '                           "t1.sekinin_nm, " & vbCrLf & _
    '                           "t1.sekinin_mail, " & vbCrLf & _
    '                           "t1.daihyo_nm, " & vbCrLf & _
    '                           "t1.riyo_tel11, " & vbCrLf & _
    '                           "t1.riyo_tel12, " & vbCrLf & _
    '                           "t1.riyo_tel13, " & vbCrLf & _
    '                           "t1.riyo_tel21, " & vbCrLf & _
    '                           "t1.riyo_tel22, " & vbCrLf & _
    '                           "t1.riyo_tel23, " & vbCrLf & _
    '                           "t1.riyo_naisen, " & vbCrLf & _
    '                           "t1.riyo_fax11, " & vbCrLf & _
    '                           "t1.riyo_fax12, " & vbCrLf & _
    '                           "t1.riyo_fax13, " & vbCrLf & _
    '                           "t1.riyo_yubin1, " & vbCrLf & _
    '                           "t1.riyo_yubin2, " & vbCrLf & _
    '                           "t1.riyo_todo, " & vbCrLf & _
    '                           "t1.riyo_shiku, " & vbCrLf & _
    '                           "t1.riyo_ban, " & vbCrLf & _
    '                           "t1.riyo_build, " & vbCrLf & _
    '                           "t1.riyo_lvl, " & vbCrLf & _
    '                           "t1.aite_cd, " & vbCrLf & _
    '                           "t1.onkyo_nm, " & vbCrLf & _
    '                           "t1.onkyo_tanto_nm, " & vbCrLf & _
    '                           "t1.onkyo_tel11, " & vbCrLf & _
    '                           "t1.onkyo_tel12, " & vbCrLf & _
    '                           "t1.onkyo_tel13, " & vbCrLf & _
    '                           "t1.onkyo_naisen, " & vbCrLf & _
    '                           "t1.onkyo_fax11, " & vbCrLf & _
    '                           "t1.onkyo_fax12, " & vbCrLf & _
    '                           "t1.onkyo_fax13, " & vbCrLf & _
    '                           "t1.onkyo_mail, " & vbCrLf & _
    '                           "t1.biko, " & vbCrLf & _
    '                           "to_char(t1.add_dt,'YYYY/MM/DD HH:mm') as add_dt, " & vbCrLf & _
    '                           "t1.add_user_cd, " & vbCrLf & _
    '                           "m1.user_nm as add_user_nm," & vbCrLf & _
    '                           "to_char(t1.up_dt,'YYYY/MM/DD HH:mm') as up_dt, " & vbCrLf & _
    '                           "t1.up_user_cd, " & vbCrLf & _
    '                           "m2.user_nm as up_user_nm, " & vbCrLf & _
    '                           "m3.aite_nm as aite_nm " & vbCrLf & _
    '                       "FROM " & vbCrLf & _
    '                           "CANCEL_WAIT_TBL t1" & vbCrLf & _
    '                       "LEFT JOIN USER_MST m1" & vbCrLf & _
    '                           "ON t1.add_user_cd = m1.user_cd" & vbCrLf & _
    '                       "LEFT JOIN USER_MST m2" & vbCrLf & _
    '                           "ON t1.up_user_cd = m2.user_cd" & vbCrLf & _
    '                       "LEFT JOIN AITESAKI_MST m3" & vbCrLf & _
    '                           "ON t1.aite_cd = m3.aite_cd" & vbCrLf & _
    '                       "WHERE" & vbCrLf & _
    '                           "cancel_wait_no = :CancelWaitNo "
    Private strEX06S001 As String = _
                            "select " & vbCrLf & _
                               "t1.cancel_wait_no, " & vbCrLf & _
                               "t1.cancel_wait_dt, " & vbCrLf & _
                               "t1.cancel_wait_usercd, " & vbCrLf & _
                               "t1.cancel_wait_sts, " & vbCrLf & _
                               "t1.shisetu_kbn, " & vbCrLf & _
                               "t1.studio_kbn, " & vbCrLf & _
                               "t1.saiji_nm, " & vbCrLf & _
                               "t1.shutsuen_nm, " & vbCrLf & _
                               "t1.kashi_kind, " & vbCrLf & _
                               "t1.riyo_type, " & vbCrLf & _
                               "t1.drink_flg, " & vbCrLf & _
                               "t1.saiji_bunrui, " & vbCrLf & _
                               "t1.teiin, " & vbCrLf & _
                               "t1.onkyo_ope_flg, " & vbCrLf & _
                               "t1.riyosha_cd, " & vbCrLf & _
                               "t1.riyo_nm, " & vbCrLf & _
                               "t1.riyo_kana, " & vbCrLf & _
                               "t1.sekinin_busho_nm, " & vbCrLf & _
                               "t1.sekinin_nm, " & vbCrLf & _
                               "t1.sekinin_mail, " & vbCrLf & _
                               "t1.daihyo_nm, " & vbCrLf & _
                               "t1.riyo_tel11, " & vbCrLf & _
                               "t1.riyo_tel12, " & vbCrLf & _
                               "t1.riyo_tel13, " & vbCrLf & _
                               "t1.riyo_tel21, " & vbCrLf & _
                               "t1.riyo_naisen, " & vbCrLf & _
                               "t1.riyo_fax11, " & vbCrLf & _
                               "t1.riyo_fax12, " & vbCrLf & _
                               "t1.riyo_fax13, " & vbCrLf & _
                               "t1.riyo_yubin1, " & vbCrLf & _
                               "t1.riyo_yubin2, " & vbCrLf & _
                               "t1.riyo_todo, " & vbCrLf & _
                               "t1.riyo_shiku, " & vbCrLf & _
                               "t1.riyo_ban, " & vbCrLf & _
                               "t1.riyo_build, " & vbCrLf & _
                               "t1.riyo_lvl, " & vbCrLf & _
                               "t1.aite_cd, " & vbCrLf & _
                               "t1.onkyo_nm, " & vbCrLf & _
                               "t1.onkyo_tanto_nm, " & vbCrLf & _
                               "t1.onkyo_tel11, " & vbCrLf & _
                               "t1.onkyo_tel12, " & vbCrLf & _
                               "t1.onkyo_tel13, " & vbCrLf & _
                               "t1.onkyo_naisen, " & vbCrLf & _
                               "t1.onkyo_fax11, " & vbCrLf & _
                               "t1.onkyo_fax12, " & vbCrLf & _
                               "t1.onkyo_fax13, " & vbCrLf & _
                               "t1.onkyo_mail, " & vbCrLf & _
                               "t1.biko, " & vbCrLf & _
                               "to_char(t1.add_dt,'YYYY/MM/DD HH:mm') as add_dt, " & vbCrLf & _
                               "t1.add_user_cd, " & vbCrLf & _
                               "m1.user_nm as add_user_nm," & vbCrLf & _
                               "to_char(t1.up_dt,'YYYY/MM/DD HH:mm') as up_dt, " & vbCrLf & _
                               "t1.up_user_cd, " & vbCrLf & _
                               "m2.user_nm as up_user_nm, " & vbCrLf & _
                               "m3.aite_nm as aite_nm " & vbCrLf & _
                           "FROM " & vbCrLf & _
                               "CANCEL_WAIT_TBL t1" & vbCrLf & _
                           "LEFT JOIN USER_MST m1" & vbCrLf & _
                               "ON t1.add_user_cd = m1.user_cd" & vbCrLf & _
                           "LEFT JOIN USER_MST m2" & vbCrLf & _
                               "ON t1.up_user_cd = m2.user_cd" & vbCrLf & _
                           "LEFT JOIN AITESAKI_MST m3" & vbCrLf & _
                               "ON t1.aite_cd = m3.aite_cd" & vbCrLf & _
                           "WHERE" & vbCrLf & _
                               "cancel_wait_no = :CancelWaitNo "
    '2016.11.4 m.hayabuchi MOD End 課題No.59

    'SQL文(キャンセル待ち日程情報の取得)
    ' 2016.01.25 UPD START↓ h.hagiwara キャンセル待ち取り消し状態の場合枠番号は取得しない
    'Private strEX06S002 As String = _
    '                       "select " & vbCrLf & _
    '                           "seq, " & vbCrLf & _
    '                           "shisetu_kbn, " & vbCrLf & _
    '                           "studio_kbn, " & vbCrLf & _
    '                           "riyo_dt_flg, " & vbCrLf & _
    '                           "riyo_ym, " & vbCrLf & _
    '                           "riyo_dt, " & vbCrLf & _
    '                           "cancel_wait_no, " & vbCrLf & _
    '                           "riyo_keitai, " & vbCrLf & _
    '                           "mitei_flg, " & vbCrLf & _
    '                           "start_time, " & vbCrLf & _
    '                           "end_time, " & vbCrLf & _
    '                           "riyo_memo, " & vbCrLf & _
    '                           "waku_no " & vbCrLf & _
    '                       "FROM " & vbCrLf & _
    '                           "CANCEL_WAIT_DT_TBL " & vbCrLf & _
    '                       "WHERE" & vbCrLf & _
    '                           "cancel_wait_no = :CancelWaitNo "
    Private strEX06S002 As String = _
                           "select " & vbCrLf & _
                               "T1.seq, " & vbCrLf & _
                               "T1.shisetu_kbn, " & vbCrLf & _
                               "T1.studio_kbn, " & vbCrLf & _
                               "T1.riyo_dt_flg, " & vbCrLf & _
                               "T1.riyo_ym, " & vbCrLf & _
                               "T1.riyo_dt, " & vbCrLf & _
                               "T1.cancel_wait_no, " & vbCrLf & _
                               "T1.riyo_keitai, " & vbCrLf & _
                               "T1.mitei_flg, " & vbCrLf & _
                               "T1.start_time, " & vbCrLf & _
                               "T1.end_time, " & vbCrLf & _
                               "T1.riyo_memo, " & vbCrLf & _
                               "CASE WHEN COALESCE(T2.cancel_wait_sts,'') = '3' " & vbCrLf & _
                               "    THEN NULL " & vbCrLf & _
                               "    ELSE T1.waku_no " & vbCrLf & _
                               "END AS waku_no " & vbCrLf & _
                           "FROM " & vbCrLf & _
                               "CANCEL_WAIT_DT_TBL T1 " & vbCrLf & _
                           "LEFT JOIN CANCEL_WAIT_TBL T2 " & vbCrLf & _
                               "ON T1.cancel_wait_no = T2.cancel_wait_no " & vbCrLf & _
                           "WHERE " & vbCrLf & _
                               "T1.cancel_wait_no = :CancelWaitNo "
    ' 2016.01.25 UPD END↑ h.hagiwara キャンセル待ち取り消し状態の場合枠番号は取得しない

    'SQL(キャンセル待ち情報登録)
    '2016.11.4 m.hayabuchi MOD Start 課題No.59
    'Private strEX06I001 As String = _
    '                       "insert into CANCEL_WAIT_TBL " & vbCrLf & _
    '                           "( " & vbCrLf & _
    '                           "CANCEL_WAIT_NO, CANCEL_WAIT_DT, CANCEL_WAIT_USERCD, CANCEL_WAIT_STS, SHISETU_KBN, STUDIO_KBN, SAIJI_NM, SHUTSUEN_NM, KASHI_KIND, RIYO_TYPE, DRINK_FLG, SAIJI_BUNRUI, TEIIN, " & vbCrLf & _
    '                               "ONKYO_OPE_FLG, RIYOSHA_CD, RIYO_NM, RIYO_KANA, SEKININ_BUSHO_NM, SEKININ_NM, SEKININ_MAIL, DAIHYO_NM, RIYO_TEL11, RIYO_TEL12, RIYO_TEL13, RIYO_TEL21, " & vbCrLf & _
    '                               "RIYO_TEL22, RIYO_TEL23, RIYO_NAISEN, RIYO_FAX11, RIYO_FAX12, RIYO_FAX13, RIYO_YUBIN1, RIYO_YUBIN2, RIYO_TODO, RIYO_SHIKU, RIYO_BAN, RIYO_BUILD, RIYO_LVL, " & vbCrLf & _
    '                               "AITE_CD, ONKYO_NM, ONKYO_TANTO_NM, ONKYO_TEL11, ONKYO_TEL12, ONKYO_TEL13, ONKYO_NAISEN, ONKYO_FAX11, ONKYO_FAX12, ONKYO_FAX13, ONKYO_MAIL, BIKO, ADD_DT, ADD_USER_CD, " & vbCrLf & _
    '                               "UP_DT, UP_USER_CD " & vbCrLf & _
    '                           ") values ( " & vbCrLf & _
    '                               "'C' || to_char(nextval('CANCEL_WAIT_TBL_S'), 'FM00000'), " & vbCrLf & _
    '                               ":CancelWaitDt, " & vbCrLf & _
    '                               ":CancelWaitUsercd, " & vbCrLf & _
    '                               ":CancelWaitSts, " & vbCrLf & _
    '                               ":ShisetuKbn, " & vbCrLf & _
    '                               ":StudioKbn, " & vbCrLf & _
    '                               ":SaijiNm, " & vbCrLf & _
    '                               ":ShutsuenNm, " & vbCrLf & _
    '                               ":KashiKind, " & vbCrLf & _
    '                               ":RiyoType, " & vbCrLf & _
    '                               ":DrinkFlg, " & vbCrLf & _
    '                               ":SaijiBunrui, " & vbCrLf & _
    '                               ":Teiin, " & vbCrLf & _
    '                               ":OnkyoOpeFlg, " & vbCrLf & _
    '                               ":RiyoshaCd, " & vbCrLf & _
    '                               ":RiyoNm, " & vbCrLf & _
    '                               ":RiyoKana, " & vbCrLf & _
    '                               ":SekininBushoNm, " & vbCrLf & _
    '                               ":SekininNm, " & vbCrLf & _
    '                               ":SekininMail, " & vbCrLf & _
    '                               ":DaihyoNm, " & vbCrLf & _
    '                               ":RiyoTel11, " & vbCrLf & _
    '                               ":RiyoTel12, " & vbCrLf & _
    '                               ":RiyoTel13, " & vbCrLf & _
    '                               ":RiyoTel21, " & vbCrLf & _
    '                               ":RiyoTel22, " & vbCrLf & _
    '                               ":RiyoTel23, " & vbCrLf & _
    '                               ":RiyoNaisen, " & vbCrLf & _
    '                               ":RiyoFax11, " & vbCrLf & _
    '                               ":RiyoFax12, " & vbCrLf & _
    '                               ":RiyoFax13, " & vbCrLf & _
    '                               ":RiyoYubin1, " & vbCrLf & _
    '                               ":RiyoYubin2, " & vbCrLf & _
    '                               ":RiyoTodo, " & vbCrLf & _
    '                               ":RiyoShiku, " & vbCrLf & _
    '                               ":RiyoBan, " & vbCrLf & _
    '                               ":RiyoBuild, " & vbCrLf & _
    '                               ":RiyoLvl, " & vbCrLf & _
    '                               ":AiteCd, " & vbCrLf & _
    '                               ":OnkyoNm, " & vbCrLf & _
    '                               ":OnkyoTantoNm, " & vbCrLf & _
    '                               ":OnkyoTel11, " & vbCrLf & _
    '                               ":OnkyoTel12, " & vbCrLf & _
    '                               ":OnkyoTel13, " & vbCrLf & _
    '                               ":OnkyoNaisen, " & vbCrLf & _
    '                               ":OnkyoFax11, " & vbCrLf & _
    '                               ":OnkyoFax12, " & vbCrLf & _
    '                               ":OnkyoFax13, " & vbCrLf & _
    '                               ":OnkyoMail, " & vbCrLf & _
    '                               ":Biko, " & vbCrLf & _
    '                               "current_timestamp, " & vbCrLf & _
    '                               ":AddUserCd, " & vbCrLf & _
    '                               "current_timestamp, " & vbCrLf & _
    '                               ":UpUserCd " & vbCrLf & _
    '                           ") RETURNING CANCEL_WAIT_NO "
    Private strEX06I001 As String = _
                           "insert into CANCEL_WAIT_TBL " & vbCrLf & _
                               "( " & vbCrLf & _
                               "CANCEL_WAIT_NO, CANCEL_WAIT_DT, CANCEL_WAIT_USERCD, CANCEL_WAIT_STS, SHISETU_KBN, STUDIO_KBN, SAIJI_NM, SHUTSUEN_NM, KASHI_KIND, RIYO_TYPE, DRINK_FLG, SAIJI_BUNRUI, TEIIN, " & vbCrLf & _
                                   "ONKYO_OPE_FLG, RIYOSHA_CD, RIYO_NM, RIYO_KANA, SEKININ_BUSHO_NM, SEKININ_NM, SEKININ_MAIL, DAIHYO_NM, RIYO_TEL11, RIYO_TEL12, RIYO_TEL13, RIYO_TEL21, " & vbCrLf & _
                                   "RIYO_NAISEN, RIYO_FAX11, RIYO_FAX12, RIYO_FAX13, RIYO_YUBIN1, RIYO_YUBIN2, RIYO_TODO, RIYO_SHIKU, RIYO_BAN, RIYO_BUILD, RIYO_LVL, " & vbCrLf & _
                                   "AITE_CD, ONKYO_NM, ONKYO_TANTO_NM, ONKYO_TEL11, ONKYO_TEL12, ONKYO_TEL13, ONKYO_NAISEN, ONKYO_FAX11, ONKYO_FAX12, ONKYO_FAX13, ONKYO_MAIL, BIKO, ADD_DT, ADD_USER_CD, " & vbCrLf & _
                                   "UP_DT, UP_USER_CD " & vbCrLf & _
                               ") values ( " & vbCrLf & _
                                   "'C' || to_char(nextval('CANCEL_WAIT_TBL_S'), 'FM00000'), " & vbCrLf & _
                                   ":CancelWaitDt, " & vbCrLf & _
                                   ":CancelWaitUsercd, " & vbCrLf & _
                                   ":CancelWaitSts, " & vbCrLf & _
                                   ":ShisetuKbn, " & vbCrLf & _
                                   ":StudioKbn, " & vbCrLf & _
                                   ":SaijiNm, " & vbCrLf & _
                                   ":ShutsuenNm, " & vbCrLf & _
                                   ":KashiKind, " & vbCrLf & _
                                   ":RiyoType, " & vbCrLf & _
                                   ":DrinkFlg, " & vbCrLf & _
                                   ":SaijiBunrui, " & vbCrLf & _
                                   ":Teiin, " & vbCrLf & _
                                   ":OnkyoOpeFlg, " & vbCrLf & _
                                   ":RiyoshaCd, " & vbCrLf & _
                                   ":RiyoNm, " & vbCrLf & _
                                   ":RiyoKana, " & vbCrLf & _
                                   ":SekininBushoNm, " & vbCrLf & _
                                   ":SekininNm, " & vbCrLf & _
                                   ":SekininMail, " & vbCrLf & _
                                   ":DaihyoNm, " & vbCrLf & _
                                   ":RiyoTel11, " & vbCrLf & _
                                   ":RiyoTel12, " & vbCrLf & _
                                   ":RiyoTel13, " & vbCrLf & _
                                   ":RiyoTel21, " & vbCrLf & _
                                   ":RiyoNaisen, " & vbCrLf & _
                                   ":RiyoFax11, " & vbCrLf & _
                                   ":RiyoFax12, " & vbCrLf & _
                                   ":RiyoFax13, " & vbCrLf & _
                                   ":RiyoYubin1, " & vbCrLf & _
                                   ":RiyoYubin2, " & vbCrLf & _
                                   ":RiyoTodo, " & vbCrLf & _
                                   ":RiyoShiku, " & vbCrLf & _
                                   ":RiyoBan, " & vbCrLf & _
                                   ":RiyoBuild, " & vbCrLf & _
                                   ":RiyoLvl, " & vbCrLf & _
                                   ":AiteCd, " & vbCrLf & _
                                   ":OnkyoNm, " & vbCrLf & _
                                   ":OnkyoTantoNm, " & vbCrLf & _
                                   ":OnkyoTel11, " & vbCrLf & _
                                   ":OnkyoTel12, " & vbCrLf & _
                                   ":OnkyoTel13, " & vbCrLf & _
                                   ":OnkyoNaisen, " & vbCrLf & _
                                   ":OnkyoFax11, " & vbCrLf & _
                                   ":OnkyoFax12, " & vbCrLf & _
                                   ":OnkyoFax13, " & vbCrLf & _
                                   ":OnkyoMail, " & vbCrLf & _
                                   ":Biko, " & vbCrLf & _
                                   "current_timestamp, " & vbCrLf & _
                                   ":AddUserCd, " & vbCrLf & _
                                   "current_timestamp, " & vbCrLf & _
                                   ":UpUserCd " & vbCrLf & _
                               ") RETURNING CANCEL_WAIT_NO "
    '2016.11.4 m.hayabuchi MOD End 課題No.59

    'SQL(キャンセル待ち情報更新)
    '2016.11.4 m.hayabuchi MOD Start 課題No.59
    'Private strEX06U001 As String = _
    '                       "update CANCEL_WAIT_TBL " & vbCrLf & _
    '                           "SET " & vbCrLf & _
    '                               "CANCEL_WAIT_DT = :CancelWaitDt, " & vbCrLf & _
    '                               "CANCEL_WAIT_USERCD = :CancelWaitUsercd, " & vbCrLf & _
    '                               "CANCEL_WAIT_STS = :CancelWaitSts, " & vbCrLf & _
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
    '                               "BIKO = :Biko, " & vbCrLf & _
    '                               "UP_DT = current_timestamp, " & vbCrLf & _
    '                               "UP_USER_CD = :UpUserCd " & vbCrLf & _
    '                           "WHERE " & vbCrLf & _
    '                               "CANCEL_WAIT_NO = :CancelWaitNo "
    Private strEX06U001 As String = _
                       "update CANCEL_WAIT_TBL " & vbCrLf & _
                           "SET " & vbCrLf & _
                               "CANCEL_WAIT_DT = :CancelWaitDt, " & vbCrLf & _
                               "CANCEL_WAIT_USERCD = :CancelWaitUsercd, " & vbCrLf & _
                               "CANCEL_WAIT_STS = :CancelWaitSts, " & vbCrLf & _
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
                               "BIKO = :Biko, " & vbCrLf & _
                               "UP_DT = current_timestamp, " & vbCrLf & _
                               "UP_USER_CD = :UpUserCd " & vbCrLf & _
                           "WHERE " & vbCrLf & _
                               "CANCEL_WAIT_NO = :CancelWaitNo "
    '2016.11.4 m.hayabuchi MOD End 課題No.59

    'SQL文(キャンセル待ち情報削除)
    Private strEX06D003 As String = _
                           "DELETE " & vbCrLf & _
                           "FROM " & vbCrLf & _
                               "CANCEL_WAIT_TBL " & vbCrLf & _
                           "WHERE" & vbCrLf & _
                                "CANCEL_WAIT_NO = :CancelWaitNo "

    'SQL文(キャンセル待ち日程情報の削除)
    Private strEX06D002 As String = _
                           "DELETE " & vbCrLf & _
                           "FROM " & vbCrLf & _
                               "CANCEL_WAIT_DT_TBL " & vbCrLf & _
                           "WHERE" & vbCrLf & _
                                "CANCEL_WAIT_NO = :CancelWaitNo "

    'SQL(キャンセル待ち日程情報の登録)
    Private strEX06I002 As String = _
                           "insert into CANCEL_WAIT_DT_TBL (SEQ, SHISETU_KBN, STUDIO_KBN, RIYO_DT_FLG, RIYO_KEITAI, MITEI_FLG, RIYO_YM, RIYO_DT, CANCEL_WAIT_NO, " & vbCrLf & _
                               "START_TIME, END_TIME, RIYO_MEMO, WAKU_NO, ADD_DT, ADD_USER_CD, UP_DT, UP_USER_CD " & vbCrLf & _
                           ")values( " & vbCrLf & _
                               "nextval('CANCEL_WAIT_DT_TBL_S'), " & vbCrLf & _
                               ":ShisetuKbn, " & vbCrLf & _
                               ":StudioKbn, " & vbCrLf & _
                               ":RiyoDtFlg, " & vbCrLf & _
                               ":RiyoKeitai, " & vbCrLf & _
                               ":MiteiFlg, " & vbCrLf & _
                               ":RiyoYm, " & vbCrLf & _
                               ":RiyoDt, " & vbCrLf & _
                               ":CancelWaitNo, " & vbCrLf & _
                               ":StartTime, " & vbCrLf & _
                               ":EndTime, " & vbCrLf & _
                               ":RiyoMemo, " & vbCrLf & _
                               ":WakuNo, " & vbCrLf & _
                               "current_timestamp, " & vbCrLf & _
                               ":AddUserCd, " & vbCrLf & _
                               "current_timestamp, " & vbCrLf & _
                               ":UpUserCd " & vbCrLf & _
                           ") "

    Private strEX99S001 As String = _
                           "SELECT " & vbCrLf & _
                           "    1 AS cnt " & vbCrLf & _
                           "FROM " & vbCrLf & _
                           "    cancel_wait_dt_tbl t1 " & vbCrLf & _
                           "    INNER JOIN " & vbCrLf & _
                           "        cancel_wait_tbl t2 " & vbCrLf & _
                           "    ON  t1.cancel_wait_no = t2.cancel_wait_no " & vbCrLf & _
                           "    AND t2.cancel_wait_sts in('1') " & vbCrLf & _
                           "WHERE " & vbCrLf & _
                           "    t1.riyo_dt = :CancelDt " & vbCrLf & _
                           "AND t1.shisetu_kbn = :ShisetuKbn " & vbCrLf & _
                           "AND t1.studio_kbn in (:StudioKbn, '3') " & vbCrLf & _
                           "AND t1.cancel_wait_no <> :CancelWaitNo " & vbCrLf & _
                           "AND t1.waku_no = :WakuNo "

    '2016.11.2 MOD START m.hayabuchi 課題No.58 ※予約ステータスが仮予約の場合は抽出しない(KAKUTEI_DTがNULLのため)
    'SQL文(責任者検索)
    'Private strEX04S004 As String = _
    '                       "SELECT" & vbCrLf & _
    '                       "    DISTINCT sekinin_nm " & vbCrLf & _
    '                       "from " & vbCrLf & _
    '                       "    yoyaku_tbl " & vbCrLf & _
    '                       "where " & vbCrLf & _
    '                       "    riyosha_cd = :RiyoshaCd " & vbCrLf & _
    '                       "AND TRIM(sekinin_nm) <> '' " & vbCrLf & _
    '                       "order by " & vbCrLf & _
    '                       "    sekinin_nm "
    Private strEX04S004 As String = _
                           "SELECT" & vbCrLf & _
                           "    DISTINCT sekinin_nm " & vbCrLf & _
                           "from " & vbCrLf & _
                           "    yoyaku_tbl " & vbCrLf & _
                           "where " & vbCrLf & _
                           "    riyosha_cd = :RiyoshaCd " & vbCrLf & _
                           "AND TRIM(sekinin_nm) <> '' " & vbCrLf & _
                           "AND yoyaku_sts in ('3','4')" & vbCrLf & _
                           "order by " & vbCrLf & _
                           "    sekinin_nm "
    '2016.11.2 MOD END m.hayabuchi 課題No.58

    '2016.11.2 MOD m.hayabuchi 課題No.59 携帯番号を一枠(RIYO_TEL21)に統合
    '2016.08.08 ADD START e.watanabe 課題No.58
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
    '2016.08.08 ADD END e.watanabe 課題No.58

    ''' <summary>
    ''' 予約情報取得
    ''' </summary>
    ''' <param name="Adapter"></param>
    ''' <param name="Cn"></param>
    ''' <param name="CancelWaitNo"></param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>ログインデータを取得するSQLの作成
    ''' <para>作成情報：2015/08/04 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function setSelectYoyakuData(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByRef Cn As NpgsqlConnection, _
                                       ByVal CancelWaitNo As String) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = String.Empty

        Try
            'SQL文(SELECT)
            'データアダプタに、SQLのSELECT文を設定
            strSQL = strEX06S001

            'Dim cmd = New OleDbCommand(strSQL, Cn)
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            '条件項目の型を指定
            'バインド変数の型を設定
            With Adapter.SelectCommand.Parameters
                .Add(New NpgsqlParameter(":CancelWaitNo", NpgsqlTypes.NpgsqlDbType.Varchar))
            End With
            'バインド変数の値を設定
            With Adapter.SelectCommand
                .Parameters(":CancelWaitNo").Value = CancelWaitNo   '予約No
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
    ''' 予約日時情報取得
    ''' </summary>
    ''' <param name="Adapter"></param>
    ''' <param name="Cn"></param>
    ''' <param name="CancelWaitNo"></param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>ログインデータを取得するSQLの作成
    ''' <para>作成情報：2015/08/04 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function setSelectRiyobiData(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByRef Cn As NpgsqlConnection, _
                                       ByVal CancelWaitNo As String) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = String.Empty

        Try
            'SQL文(SELECT)
            'データアダプタに、SQLのSELECT文を設定
            strSQL = strEX06S002

            'Dim cmd = New OleDbCommand(strSQL, Cn)
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            '条件項目の型を指定
            'バインド変数の型を設定
            With Adapter.SelectCommand.Parameters
                .Add(New NpgsqlParameter(":CancelWaitNo", NpgsqlTypes.NpgsqlDbType.Varchar))
            End With
            'バインド変数の値を設定
            With Adapter.SelectCommand
                .Parameters(":CancelWaitNo").Value = CancelWaitNo   '予約No
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
    ''' 予約情報テーブルの登録
    ''' </summary>
    ''' <param name="Cmd"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataEXTC0104"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' <para>作成情報：2015/08/05 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function registerYoyakuInfo(ByRef Cmd As NpgsqlCommand, _
                                   ByRef Cn As NpgsqlConnection, _
                                   ByVal dataEXTC0104 As DataEXTC0104, _
                                   ByVal blnIsUpdate As Boolean) As Boolean

        Dim strSQL As String = String.Empty
        Try
            'SQL文を設定
            If blnIsUpdate = True Then
                strSQL = strEX06U001
            Else
                strSQL = strEX06I001
            End If
            Cmd = New NpgsqlCommand(strSQL, Cn)

            '値を設定
            Cmd.Parameters.Add(New NpgsqlParameter(":CancelWaitDt", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":CancelWaitUsercd", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":CancelWaitSts", NpgsqlTypes.NpgsqlDbType.Char))
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
            Cmd.Parameters.Add(New NpgsqlParameter(":RiyoLvl", NpgsqlTypes.NpgsqlDbType.Varchar))
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
            Cmd.Parameters.Add(New NpgsqlParameter(":Biko", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":AddUserCd", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":UpUserCd", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":CancelWaitNo", NpgsqlTypes.NpgsqlDbType.Varchar))

            '2016.11.4 m.hayabuchi MOD Start 課題No.59
            'Cmd.Parameters(0).Value = dataEXTC0104.PropStrCanUkeDt
            'Cmd.Parameters(1).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrCanUkeUsercd)
            'Cmd.Parameters(2).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrYoyakuSts)
            'Cmd.Parameters(3).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrShisetuKbn)
            'Cmd.Parameters(4).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrStudioKbn)
            'Cmd.Parameters(5).Value = DBNull.Value
            'Cmd.Parameters(6).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrShutsuenNm)
            'Cmd.Parameters(7).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrKashiKind)
            'Cmd.Parameters(8).Value = DBNull.Value
            'Cmd.Parameters(9).Value = DBNull.Value
            'Cmd.Parameters(10).Value = DBNull.Value
            'Cmd.Parameters(11).Value = DBNull.Value
            'Cmd.Parameters(12).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrOnkyoOpe)
            'Cmd.Parameters(13).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrRiyoshaCd)
            'Cmd.Parameters(14).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrRiyoNm)
            'Cmd.Parameters(15).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrRiyoKana)
            'Cmd.Parameters(16).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrSekininBushoNm)
            'Cmd.Parameters(17).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrSekininNm)
            'Cmd.Parameters(18).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrSekininMail)
            'Cmd.Parameters(19).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrDaihyoNm)
            'Cmd.Parameters(20).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrRiyoTel11)
            'Cmd.Parameters(21).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrRiyoTel12)
            'Cmd.Parameters(22).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrRiyoTel13)
            'Cmd.Parameters(23).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrRiyoTel21)
            'Cmd.Parameters(24).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrRiyoTel22)
            'Cmd.Parameters(25).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrRiyoTel23)
            'Cmd.Parameters(26).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrRiyoNaisen)
            'Cmd.Parameters(27).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrRiyoFax11)
            'Cmd.Parameters(28).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrRiyoFax12)
            'Cmd.Parameters(29).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrRiyoFax13)
            'Cmd.Parameters(30).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrRiyoYubin1)
            'Cmd.Parameters(31).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrRiyoYubin2)
            'Cmd.Parameters(32).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrRiyoTodo)
            'Cmd.Parameters(33).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrRiyoShiku)
            'Cmd.Parameters(34).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrRiyoBan)
            'Cmd.Parameters(35).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrRiyoBuild)
            'Cmd.Parameters(36).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrRiyoLvl)
            'Cmd.Parameters(37).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrAiteCd)
            'Cmd.Parameters(38).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrOnkyoNm)
            'Cmd.Parameters(39).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrOnkyoTantoNm)
            'Cmd.Parameters(40).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrOnkyoTel11)
            'Cmd.Parameters(41).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrOnkyoTel12)
            'Cmd.Parameters(42).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrOnkyoTel13)
            'Cmd.Parameters(43).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrOnkyoNaisen)
            'Cmd.Parameters(44).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrOnkyoFax11)
            'Cmd.Parameters(45).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrOnkyoFax12)
            'Cmd.Parameters(46).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrOnkyoFax13)
            'Cmd.Parameters(47).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrOnkyoMail)
            'Cmd.Parameters(48).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrBiko)
            'Cmd.Parameters(49).Value = CommonEXT.PropComStrUserId
            'Cmd.Parameters(50).Value = CommonEXT.PropComStrUserId
            'Cmd.Parameters(51).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrYoyakuNo)
            Cmd.Parameters(0).Value = dataEXTC0104.PropStrCanUkeDt
            Cmd.Parameters(1).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrCanUkeUsercd)
            Cmd.Parameters(2).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrYoyakuSts)
            Cmd.Parameters(3).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrShisetuKbn)
            Cmd.Parameters(4).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrStudioKbn)
            Cmd.Parameters(5).Value = DBNull.Value
            Cmd.Parameters(6).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrShutsuenNm)
            Cmd.Parameters(7).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrKashiKind)
            Cmd.Parameters(8).Value = DBNull.Value
            Cmd.Parameters(9).Value = DBNull.Value
            Cmd.Parameters(10).Value = DBNull.Value
            Cmd.Parameters(11).Value = DBNull.Value
            Cmd.Parameters(12).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrOnkyoOpe)
            Cmd.Parameters(13).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrRiyoshaCd)
            Cmd.Parameters(14).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrRiyoNm)
            Cmd.Parameters(15).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrRiyoKana)
            Cmd.Parameters(16).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrSekininBushoNm)
            Cmd.Parameters(17).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrSekininNm)
            Cmd.Parameters(18).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrSekininMail)
            Cmd.Parameters(19).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrDaihyoNm)
            Cmd.Parameters(20).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrRiyoTel11)
            Cmd.Parameters(21).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrRiyoTel12)
            Cmd.Parameters(22).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrRiyoTel13)
            Cmd.Parameters(23).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrRiyoTel21)
            Cmd.Parameters(24).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrRiyoNaisen)
            Cmd.Parameters(25).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrRiyoFax11)
            Cmd.Parameters(26).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrRiyoFax12)
            Cmd.Parameters(27).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrRiyoFax13)
            Cmd.Parameters(28).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrRiyoYubin1)
            Cmd.Parameters(29).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrRiyoYubin2)
            Cmd.Parameters(30).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrRiyoTodo)
            Cmd.Parameters(31).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrRiyoShiku)
            Cmd.Parameters(32).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrRiyoBan)
            Cmd.Parameters(33).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrRiyoBuild)
            Cmd.Parameters(34).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrRiyoLvl)
            Cmd.Parameters(35).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrAiteCd)
            Cmd.Parameters(36).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrOnkyoNm)
            Cmd.Parameters(37).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrOnkyoTantoNm)
            Cmd.Parameters(38).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrOnkyoTel11)
            Cmd.Parameters(39).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrOnkyoTel12)
            Cmd.Parameters(40).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrOnkyoTel13)
            Cmd.Parameters(41).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrOnkyoNaisen)
            Cmd.Parameters(42).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrOnkyoFax11)
            Cmd.Parameters(43).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrOnkyoFax12)
            Cmd.Parameters(44).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrOnkyoFax13)
            Cmd.Parameters(45).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrOnkyoMail)
            Cmd.Parameters(46).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrBiko)
            Cmd.Parameters(47).Value = CommonEXT.PropComStrUserId
            Cmd.Parameters(48).Value = CommonEXT.PropComStrUserId
            Cmd.Parameters(49).Value = commonLogicEXT.convInsStr(dataEXTC0104.PropStrYoyakuNo)
            '2016.11.4 m.hayabuchi MOD End 課題No.59

            '正常終了
            Return True
        Catch ex As Exception
            '例外処理
            Common.CommonDeclare.puErrMsg = ex.Message
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            Return False
        End Try

    End Function

    ''' <summary>
    ''' 予約日程表テーブルのデータ削除
    ''' </summary>
    ''' <param name="Cmd"></param>
    ''' <param name="Cn"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' <para>作成情報：2015/08/05 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function deleteYoyakuList(ByRef Cmd As NpgsqlCommand, _
                                   ByRef Cn As NpgsqlConnection, _
                                   ByVal CancelWaitNo As String) As Boolean

        Dim strSQL As String = String.Empty
        Try
            'SQL文を設定
            strSQL = strEX06D002
            Cmd = New NpgsqlCommand(strSQL, Cn)

            '値を設定
            Cmd.Parameters.Add(New NpgsqlParameter(":CancelWaitNo", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters(0).Value = CancelWaitNo      '予約No

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
    ''' 予約日程表テーブルの登録
    ''' </summary>
    ''' <param name="Cmd"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataRiyobi"></param>
    ''' <param name="CancelWaitNo"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' <para>作成情報：2015/08/05 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function insertYoyakuList(ByRef Cmd As NpgsqlCommand, _
                                   ByRef Cn As NpgsqlConnection, _
                                   ByVal dataRiyobi As CommonDataCancel, _
                                   ByVal CancelWaitNo As String) As Boolean

        Dim strSQL As String = String.Empty
        Try
            'SQL文を設定
            strSQL = strEX06I002
            Cmd = New NpgsqlCommand(strSQL, Cn)

            '値を設定
            Cmd.Parameters.Add(New NpgsqlParameter(":ShisetuKbn", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":StudioKbn", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":RiyoDtFlg", NpgsqlTypes.NpgsqlDbType.Char))
            Cmd.Parameters.Add(New NpgsqlParameter(":RiyoKeitai", NpgsqlTypes.NpgsqlDbType.Char))
            Cmd.Parameters.Add(New NpgsqlParameter(":MiteiFlg", NpgsqlTypes.NpgsqlDbType.Char))
            Cmd.Parameters.Add(New NpgsqlParameter(":RiyoYm", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":RiyoDt", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":CancelWaitNo", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":StartTime", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":EndTime", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":RiyoMemo", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":WakuNo", NpgsqlTypes.NpgsqlDbType.Integer))
            Cmd.Parameters.Add(New NpgsqlParameter(":AddUserCd", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":UpUserCd", NpgsqlTypes.NpgsqlDbType.Varchar))

            Cmd.Parameters(0).Value = dataRiyobi.PropStrShisetuKbn
            Cmd.Parameters(1).Value = dataRiyobi.PropStrStudioKbn
            Cmd.Parameters(2).Value = dataRiyobi.PropStrKibobiKbn
            Cmd.Parameters(3).Value = dataRiyobi.PropStrRiyoKeitai
            Cmd.Parameters(4).Value = dataRiyobi.PropStrMiteiFlg
            Cmd.Parameters(5).Value = commonLogicEXT.convInsStr(dataRiyobi.PropStrCancelYm)
            Cmd.Parameters(6).Value = commonLogicEXT.convInsStr(dataRiyobi.PropStrCancelDt)
            Cmd.Parameters(7).Value = CancelWaitNo
            Cmd.Parameters(8).Value = commonLogicEXT.convInsStr(dataRiyobi.PropStrStartTime)
            Cmd.Parameters(9).Value = commonLogicEXT.convInsStr(dataRiyobi.PropStrEndTime)
            Cmd.Parameters(10).Value = commonLogicEXT.convInsStr(dataRiyobi.PropStrMemo)
            Cmd.Parameters(11).Value = commonLogicEXT.convInsNum(dataRiyobi.PropIntWakuNo)
            Cmd.Parameters(12).Value = CommonEXT.PropComStrUserId
            Cmd.Parameters(13).Value = CommonEXT.PropComStrUserId

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
    ''' 予約テーブルのデータ削除
    ''' </summary>
    ''' <param name="Cmd"></param>
    ''' <param name="Cn"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' <para>作成情報：2015/08/17 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function deleteYoyaku(ByRef Cmd As NpgsqlCommand, _
                                   ByRef Cn As NpgsqlConnection, _
                                   ByVal CancelWaitNo As String) As Boolean

        Dim strSQL As String = String.Empty
        Try
            'SQL文を設定
            strSQL = strEX06D003
            Cmd = New NpgsqlCommand(strSQL, Cn)

            '値を設定
            Cmd.Parameters.Add(New NpgsqlParameter(":CancelWaitNo", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters(0).Value = CancelWaitNo      '予約No

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
    ''' キャンセル待ち日時情報取得
    ''' </summary>
    ''' <param name="Adapter"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataRiyobi"></param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>ログインデータを取得するSQLの作成
    ''' <para>作成情報：2015/08/04 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function selectWakuNo(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByRef Cn As NpgsqlConnection, _
                                       ByVal dataRiyobi As CommonDataCancel, _
                                       ByVal wakuNo As Integer) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = String.Empty
        Dim Table As New DataTable()

        Try
            'SQL文(SELECT)
            'データアダプタに、SQLのSELECT文を設定
            strSQL = strEX99S001

            'Dim cmd = New OleDbCommand(strSQL, Cn)
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            '条件項目の型を指定
            'バインド変数の型を設定
            With Adapter.SelectCommand.Parameters
                .Add(New NpgsqlParameter(":CancelWaitNo", NpgsqlTypes.NpgsqlDbType.Varchar))
                .Add(New NpgsqlParameter(":CancelDt", NpgsqlTypes.NpgsqlDbType.Varchar))
                .Add(New NpgsqlParameter(":ShisetuKbn", NpgsqlTypes.NpgsqlDbType.Char))
                .Add(New NpgsqlParameter(":StudioKbn", NpgsqlTypes.NpgsqlDbType.Char))
                .Add(New NpgsqlParameter(":WakuNo", NpgsqlTypes.NpgsqlDbType.Integer))
            End With
            'バインド変数の値を設定
            With Adapter.SelectCommand
                If dataRiyobi.PropStrCancelWaitNo Is Nothing Then
                    .Parameters(0).Value = ""
                Else
                    .Parameters(0).Value = dataRiyobi.PropStrCancelWaitNo
                End If
                .Parameters(1).Value = dataRiyobi.PropStrCancelDt
                .Parameters(2).Value = dataRiyobi.PropStrShisetuKbn
                .Parameters(3).Value = dataRiyobi.PropStrStudioKbn
                .Parameters(4).Value = wakuNo
            End With
            Adapter.Fill(Table)
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "GetWaku", Nothing, Adapter.SelectCommand)
            If Table.Rows.Count > 0 Then
                Return False
            End If
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True
        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            Return False
        Finally
            Table.Dispose()
        End Try

    End Function

    ''' <summary>
    ''' 責任者情報取得
    ''' </summary>
    ''' <param name="Adapter"></param>
    ''' <param name="Cn"></param>
    ''' <param name="strRiyoshaCd"></param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>ログインデータを取得するSQLの作成
    ''' <para>作成情報：2015/08/04 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function setSelectSekininName(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByRef Cn As NpgsqlConnection, _
                                       ByVal strRiyoshaCd As String) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = String.Empty

        Try
            'SQL文(SELECT)
            'データアダプタに、SQLのSELECT文を設定
            strSQL = strEX04S004
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            '条件項目の型を指定
            'バインド変数の型を設定
            With Adapter.SelectCommand.Parameters
                .Add(New NpgsqlParameter(":RiyoshaCd", NpgsqlTypes.NpgsqlDbType.Varchar))
            End With
            'バインド変数の値を設定
            With Adapter.SelectCommand
                .Parameters(":RiyoshaCd").Value = strRiyoshaCd
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
    ''' 責任者メールアドレス・携帯電話番号情報取得
    ''' </summary>
    ''' <param name="Adapter"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataEXTC0104"></param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>責任者名から責任者メールアドレス・携帯電話番号を取得するSQL作成
    ''' <para>作成情報：2016/11/4 m.hayabuchi
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function GetSqlSekininshaMailTelData(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByRef Cn As NpgsqlConnection, _
                                       ByVal dataEXTC0104 As DataEXTC0104) As Boolean

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
                .Parameters(":SekininNm").Value = dataEXTC0104.PropStrSekininNm   '責任者名
                .Parameters(":RiyoshaCd").Value = dataEXTC0104.PropStrRiyoshaCd   '利用者コード 
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

End Class
