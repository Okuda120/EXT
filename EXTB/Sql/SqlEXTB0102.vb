Imports Common
Imports Npgsql
Imports CommonEXT
Imports EXTZ

Public Class SqlEXTB0102

    Private commonLogicEXT As New CommonLogicEXT

    'SQL文(ログイン)
    '2016.11.4 m.hayabuchi MOD Start 課題No.59
    'Private strEX04S001 As String = _
    '                        "select " & vbCrLf & _
    '                           "t1.yoyaku_no, " & vbCrLf & _
    '                           "t1.kariuke_dt, " & vbCrLf & _
    '                           "t1.kari_usercd, " & vbCrLf & _
    '                           "t1.kakutei_dt, " & vbCrLf & _
    '                           "t1.kaku_usercd, " & vbCrLf & _
    '                           "t1.yoyaku_sts, " & vbCrLf & _
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
    '                           "t1.total_riyo_kin, " & vbCrLf & _
    '                           "t1.send_kbn, " & vbCrLf & _
    '                           "t1.send_sts, " & vbCrLf & _
    '                           "t1.send_dt, " & vbCrLf & _
    '                           "t1.henso_dt, " & vbCrLf & _
    '                           "t1.biko, " & vbCrLf & _
    '                           "t1.riyo_com, " & vbCrLf & _
    '                           "t1.ticket_enter_kbn, " & vbCrLf & _
    '                           "t1.ticket_drink_kbn, " & vbCrLf & _
    '                           "t1.hp_keisai, " & vbCrLf & _
    '                           "t1.joho_kokai_dt, " & vbCrLf & _
    '                           "t1.joho_kokai_time, " & vbCrLf & _
    '                           "t1.kokai_dt, " & vbCrLf & _
    '                           "t1.kokai_time, " & vbCrLf & _
    '                           "t1.finput_sts, " & vbCrLf & _
    '                           "to_char(t1.add_dt,'YYYY/MM/DD HH:mm') as add_dt, " & vbCrLf & _
    '                           "t1.add_user_cd, " & vbCrLf & _
    '                           "m1.user_nm as add_user_nm," & vbCrLf & _
    '                           "to_char(t1.up_dt,'YYYY/MM/DD HH:mm') as up_dt, " & vbCrLf & _
    '                           "t1.up_user_cd, " & vbCrLf & _
    '                           "m2.user_nm as up_user_nm, " & vbCrLf & _
    '                           "m3.aite_nm as aite_nm " & vbCrLf & _
    '                       "FROM " & vbCrLf & _
    '                           "YOYAKU_TBL t1" & vbCrLf & _
    '                       "LEFT JOIN USER_MST m1" & vbCrLf & _
    '                           "ON t1.add_user_cd = m1.user_cd" & vbCrLf & _
    '                       "LEFT JOIN USER_MST m2" & vbCrLf & _
    '                           "ON t1.up_user_cd = m2.user_cd" & vbCrLf & _
    '                       "LEFT JOIN AITESAKI_MST m3" & vbCrLf & _
    '                           "ON t1.aite_cd = m3.aite_cd" & vbCrLf & _
    '                       "WHERE" & vbCrLf & _
    '                           "YOYAKU_NO = :YoyakuNo "
    Private strEX04S001 As String = _
                        "select " & vbCrLf & _
                           "t1.yoyaku_no, " & vbCrLf & _
                           "t1.kariuke_dt, " & vbCrLf & _
                           "t1.kari_usercd, " & vbCrLf & _
                           "t1.kakutei_dt, " & vbCrLf & _
                           "t1.kaku_usercd, " & vbCrLf & _
                           "t1.yoyaku_sts, " & vbCrLf & _
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
                           "t1.total_riyo_kin, " & vbCrLf & _
                           "t1.send_kbn, " & vbCrLf & _
                           "t1.send_sts, " & vbCrLf & _
                           "t1.send_dt, " & vbCrLf & _
                           "t1.henso_dt, " & vbCrLf & _
                           "t1.biko, " & vbCrLf & _
                           "t1.riyo_com, " & vbCrLf & _
                           "t1.ticket_enter_kbn, " & vbCrLf & _
                           "t1.ticket_drink_kbn, " & vbCrLf & _
                           "t1.hp_keisai, " & vbCrLf & _
                           "t1.joho_kokai_dt, " & vbCrLf & _
                           "t1.joho_kokai_time, " & vbCrLf & _
                           "t1.kokai_dt, " & vbCrLf & _
                           "t1.kokai_time, " & vbCrLf & _
                           "t1.finput_sts, " & vbCrLf & _
                           "to_char(t1.add_dt,'YYYY/MM/DD HH:mm') as add_dt, " & vbCrLf & _
                           "t1.add_user_cd, " & vbCrLf & _
                           "m1.user_nm as add_user_nm," & vbCrLf & _
                           "to_char(t1.up_dt,'YYYY/MM/DD HH:mm') as up_dt, " & vbCrLf & _
                           "t1.up_user_cd, " & vbCrLf & _
                           "m2.user_nm as up_user_nm, " & vbCrLf & _
                           "m3.aite_nm as aite_nm " & vbCrLf & _
                       "FROM " & vbCrLf & _
                           "YOYAKU_TBL t1" & vbCrLf & _
                       "LEFT JOIN USER_MST m1" & vbCrLf & _
                           "ON t1.add_user_cd = m1.user_cd" & vbCrLf & _
                       "LEFT JOIN USER_MST m2" & vbCrLf & _
                           "ON t1.up_user_cd = m2.user_cd" & vbCrLf & _
                       "LEFT JOIN AITESAKI_MST m3" & vbCrLf & _
                           "ON t1.aite_cd = m3.aite_cd" & vbCrLf & _
                       "WHERE" & vbCrLf & _
                           "YOYAKU_NO = :YoyakuNo "
    '2016.11.4 m.hayabuchi MOD End 課題No.59

    'SQL文(ログイン)
    Private strEX04S002 As String = _
                           "select " & vbCrLf & _
                               "seq, " & vbCrLf & _
                               "shisetu_kbn, " & vbCrLf & _
                               "studio_kbn, " & vbCrLf & _
                               "yoyaku_dt, " & vbCrLf & _
                               "start_time, " & vbCrLf & _
                               "end_time, " & vbCrLf & _
                               "yoyaku_no, " & vbCrLf & _
                               "riyo_keitai, " & vbCrLf & _
                               "mitei_flg, " & vbCrLf & _
                               "tanka, " & vbCrLf & _
                               "bairitu, " & vbCrLf & _
                               "su, " & vbCrLf & _
                               "riyo_kin " & vbCrLf & _
                           "FROM " & vbCrLf & _
                               "YDT_TBL " & vbCrLf & _
                           "WHERE" & vbCrLf & _
                               "yoyaku_no = :YoyakuNo "

    'SQL(予約情報登録)
    '2016.11.4 m.hayabuchi MOD Start 課題No.59
    'Private strEX04I001 As String = _
    '                       "insert into YOYAKU_TBL " & vbCrLf & _
    '                           "( " & vbCrLf & _
    '                           "YOYAKU_NO, KARIUKE_DT, KARI_USERCD, KAKUTEI_DT, KAKU_USERCD, YOYAKU_STS, SHISETU_KBN, " & vbCrLf & _
    '                               "STUDIO_KBN, SAIJI_NM, SHUTSUEN_NM, KASHI_KIND, RIYO_TYPE, DRINK_FLG, SAIJI_BUNRUI, TEIIN, ONKYO_OPE_FLG," & vbCrLf & _
    '                               "RIYOSHA_CD, RIYO_NM, RIYO_KANA, SEKININ_BUSHO_NM, SEKININ_NM, SEKININ_MAIL, DAIHYO_NM, RIYO_TEL11, " & vbCrLf & _
    '                               "RIYO_TEL12, RIYO_TEL13, RIYO_TEL21, RIYO_TEL22, RIYO_TEL23, RIYO_NAISEN, RIYO_FAX11, RIYO_FAX12, " & vbCrLf & _
    '                               "RIYO_FAX13, RIYO_YUBIN1, RIYO_YUBIN2, RIYO_TODO, RIYO_SHIKU, RIYO_BAN, RIYO_BUILD, RIYO_LVL, AITE_CD, " & vbCrLf & _
    '                               "ONKYO_NM, ONKYO_TANTO_NM, ONKYO_TEL11, ONKYO_TEL12, ONKYO_TEL13, ONKYO_NAISEN, ONKYO_FAX11, ONKYO_FAX12, " & vbCrLf & _
    '                               "ONKYO_FAX13, ONKYO_MAIL, TOTAL_RIYO_KIN, SEND_KBN, SEND_STS, SEND_DT, HENSO_DT, BIKO, RIYO_COM, " & vbCrLf & _
    '                               "TICKET_ENTER_KBN, TICKET_DRINK_KBN, HP_KEISAI, JOHO_KOKAI_DT, JOHO_KOKAI_TIME, KOKAI_DT, KOKAI_TIME, " & vbCrLf & _
    '                               "FINPUT_STS, ADD_DT, ADD_USER_CD, UP_DT, UP_USER_CD " & vbCrLf & _
    '                           ") values ( " & vbCrLf & _
    '                               "'T' || to_char(nextval('YOYAKU_TBL_S1'), 'FM00000'), " & vbCrLf & _
    '                               ":KariukeDt, " & vbCrLf & _
    '                               ":KariUsercd, " & vbCrLf & _
    '                               ":KakuteiDt, " & vbCrLf & _
    '                               ":KakuUsercd, " & vbCrLf & _
    '                               ":YoyakuSts, " & vbCrLf & _
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
    '                               ":TotalRiyoKin, " & vbCrLf & _
    '                               ":SendKbn, " & vbCrLf & _
    '                               ":SendSts, " & vbCrLf & _
    '                               ":SendDt, " & vbCrLf & _
    '                               ":HensoDt, " & vbCrLf & _
    '                               ":Biko, " & vbCrLf & _
    '                               ":RiyoCom, " & vbCrLf & _
    '                               ":TicketEnterKbn, " & vbCrLf & _
    '                               ":TicketDrinkKbn, " & vbCrLf & _
    '                               ":HpKeisai, " & vbCrLf & _
    '                               ":JohoKokaiDt, " & vbCrLf & _
    '                               ":JohoKokaiTime, " & vbCrLf & _
    '                               ":KokaiDt, " & vbCrLf & _
    '                               ":KokaiTime, " & vbCrLf & _
    '                               ":FinputSts, " & vbCrLf & _
    '                               "current_timestamp, " & vbCrLf & _
    '                               ":AddUserCd, " & vbCrLf & _
    '                               "current_timestamp, " & vbCrLf & _
    '                               ":UpUserCd " & vbCrLf & _
    '                           ") RETURNING YOYAKU_NO "

    Private strEX04I001 As String = _
                       "insert into YOYAKU_TBL " & vbCrLf & _
                           "( " & vbCrLf & _
                           "YOYAKU_NO, KARIUKE_DT, KARI_USERCD, KAKUTEI_DT, KAKU_USERCD, YOYAKU_STS, SHISETU_KBN, " & vbCrLf & _
                               "STUDIO_KBN, SAIJI_NM, SHUTSUEN_NM, KASHI_KIND, RIYO_TYPE, DRINK_FLG, SAIJI_BUNRUI, TEIIN, ONKYO_OPE_FLG," & vbCrLf & _
                               "RIYOSHA_CD, RIYO_NM, RIYO_KANA, SEKININ_BUSHO_NM, SEKININ_NM, SEKININ_MAIL, DAIHYO_NM, RIYO_TEL11, " & vbCrLf & _
                               "RIYO_TEL12, RIYO_TEL13, RIYO_TEL21, RIYO_NAISEN, RIYO_FAX11, RIYO_FAX12, " & vbCrLf & _
                               "RIYO_FAX13, RIYO_YUBIN1, RIYO_YUBIN2, RIYO_TODO, RIYO_SHIKU, RIYO_BAN, RIYO_BUILD, RIYO_LVL, AITE_CD, " & vbCrLf & _
                               "ONKYO_NM, ONKYO_TANTO_NM, ONKYO_TEL11, ONKYO_TEL12, ONKYO_TEL13, ONKYO_NAISEN, ONKYO_FAX11, ONKYO_FAX12, " & vbCrLf & _
                               "ONKYO_FAX13, ONKYO_MAIL, TOTAL_RIYO_KIN, SEND_KBN, SEND_STS, SEND_DT, HENSO_DT, BIKO, RIYO_COM, " & vbCrLf & _
                               "TICKET_ENTER_KBN, TICKET_DRINK_KBN, HP_KEISAI, JOHO_KOKAI_DT, JOHO_KOKAI_TIME, KOKAI_DT, KOKAI_TIME, " & vbCrLf & _
                               "FINPUT_STS, ADD_DT, ADD_USER_CD, UP_DT, UP_USER_CD " & vbCrLf & _
                           ") values ( " & vbCrLf & _
                               "'T' || to_char(nextval('YOYAKU_TBL_S1'), 'FM00000'), " & vbCrLf & _
                               ":KariukeDt, " & vbCrLf & _
                               ":KariUsercd, " & vbCrLf & _
                               ":KakuteiDt, " & vbCrLf & _
                               ":KakuUsercd, " & vbCrLf & _
                               ":YoyakuSts, " & vbCrLf & _
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
                               ":TotalRiyoKin, " & vbCrLf & _
                               ":SendKbn, " & vbCrLf & _
                               ":SendSts, " & vbCrLf & _
                               ":SendDt, " & vbCrLf & _
                               ":HensoDt, " & vbCrLf & _
                               ":Biko, " & vbCrLf & _
                               ":RiyoCom, " & vbCrLf & _
                               ":TicketEnterKbn, " & vbCrLf & _
                               ":TicketDrinkKbn, " & vbCrLf & _
                               ":HpKeisai, " & vbCrLf & _
                               ":JohoKokaiDt, " & vbCrLf & _
                               ":JohoKokaiTime, " & vbCrLf & _
                               ":KokaiDt, " & vbCrLf & _
                               ":KokaiTime, " & vbCrLf & _
                               ":FinputSts, " & vbCrLf & _
                               "current_timestamp, " & vbCrLf & _
                               ":AddUserCd, " & vbCrLf & _
                               "current_timestamp, " & vbCrLf & _
                               ":UpUserCd " & vbCrLf & _
                           ") RETURNING YOYAKU_NO "
    '2016.11.4 m.hayabuchi MOD End 課題No.59

    'SQL(予約情報更新)
    '2016.11.4 m.hayabuchi MOD Start 課題No.59
    'Private strEX04U001 As String = _
    '                       "update YOYAKU_TBL " & vbCrLf & _
    '                           "SET " & vbCrLf & _
    '                               "KARIUKE_DT = :KariukeDt, " & vbCrLf & _
    '                               "KARI_USERCD = :KariUsercd, " & vbCrLf & _
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
    '                               "SEND_KBN = :SendKbn, " & vbCrLf & _
    '                               "SEND_STS = :SendSts, " & vbCrLf & _
    '                               "SEND_DT = :SendDt, " & vbCrLf & _
    '                               "HENSO_DT = :HensoDt, " & vbCrLf & _
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
    'SQL(予約情報更新)
    Private strEX04U001 As String = _
                           "update YOYAKU_TBL " & vbCrLf & _
                               "SET " & vbCrLf & _
                                   "KARIUKE_DT = :KariukeDt, " & vbCrLf & _
                                   "KARI_USERCD = :KariUsercd, " & vbCrLf & _
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
                                   "SEND_KBN = :SendKbn, " & vbCrLf & _
                                   "SEND_STS = :SendSts, " & vbCrLf & _
                                   "SEND_DT = :SendDt, " & vbCrLf & _
                                   "HENSO_DT = :HensoDt, " & vbCrLf & _
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
    '2016.11.4 m.hayabuchi MOD End 課題No.59

    'SQL文(予約データ削除)
    Private strEX04D001 As String = _
                           "DELETE " & vbCrLf & _
                           "FROM " & vbCrLf & _
                               "YOYAKU_TBL " & vbCrLf & _
                           "WHERE" & vbCrLf & _
                                "YOYAKU_NO = :YoyakuNo "

    'SQL文(予約受付制御データ削除)
    Private strEX04D002 As String = _
                           "DELETE " & vbCrLf & _
                           "FROM " & vbCrLf & _
                               "YDT_TBL " & vbCrLf & _
                           "WHERE" & vbCrLf & _
                                "YOYAKU_NO = :YoyakuNo "

    'SQL(予約日程表登録)
    Private strEX04I003 As String = _
                           "insert into YDT_TBL(SEQ, SHISETU_KBN, STUDIO_KBN, YOYAKU_DT, START_TIME, END_TIME, YOYAKU_NO, RIYO_KEITAI, " & vbCrLf & _
                               "MITEI_FLG, TANKA, BAIRITU, SU, RIYO_KIN, ADD_DT, ADD_USER_CD, UP_DT, UP_USER_CD) " & vbCrLf & _
                           "values( " & vbCrLf & _
                               "nextval('YDT_TBL_S'), " & vbCrLf & _
                               ":ShisetuKbn, " & vbCrLf & _
                               ":StudioKbn, " & vbCrLf & _
                               ":YoyakuDt, " & vbCrLf & _
                               ":StartTime, " & vbCrLf & _
                               ":EndTime, " & vbCrLf & _
                               ":YoyakuNo, " & vbCrLf & _
                               ":RiyoKeitai, " & vbCrLf & _
                               ":MiteiFlg, " & vbCrLf & _
                               ":Tanka, " & vbCrLf & _
                               ":Bairitu, " & vbCrLf & _
                               ":Su, " & vbCrLf & _
                               ":RiyoKin, " & vbCrLf & _
                               "current_timestamp, " & vbCrLf & _
                               ":AddUserCd, " & vbCrLf & _
                               "current_timestamp, " & vbCrLf & _
                               ":UpUserCd " & vbCrLf & _
                           ") "

    'SQL文(利用者チェック)
    Private strEX04S003 As String = _
                           "select " & vbCrLf & _
                               "RIYOSHA_CD " & vbCrLf & _
                           "FROM " & vbCrLf & _
                               "RIYOSHA_MST " & vbCrLf & _
                           "WHERE" & vbCrLf & _
                               "RIYO_KANA = :RiyoshaNm "

    '2016.11.2 MOD START m.hayabuchi 課題No.58 ※予約ステータスが仮予約の場合は抽出しない(KAKUTEI_DTがNULLのため)
    'SQL文(責任者検索)
    'Private strEX04S004 As String = _
    '                      "SELECT" & vbCrLf & _
    '                      "    DISTINCT sekinin_nm " & vbCrLf & _
    '                      "from " & vbCrLf & _
    '                      "    yoyaku_tbl " & vbCrLf & _
    '                      "where " & vbCrLf & _
    '                      "    riyosha_cd = :RiyoshaCd " & vbCrLf & _
    '                      "AND TRIM(sekinin_nm) <> '' " & vbCrLf & _
    '                      "order by " & vbCrLf & _
    '                      "    sekinin_nm "
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

    ''' <summary>
    ''' 予約情報取得
    ''' </summary>
    ''' <param name="Adapter"></param>
    ''' <param name="Cn"></param>
    ''' <param name="YoyakuNo"></param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>ログインデータを取得するSQLの作成
    ''' <para>作成情報：2015/08/04 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function setSelectYoyakuData(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByRef Cn As NpgsqlConnection, _
                                       ByVal YoyakuNo As String) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = String.Empty

        Try
            'SQL文(SELECT)
            'データアダプタに、SQLのSELECT文を設定
            strSQL = strEX04S001

            'Dim cmd = New OleDbCommand(strSQL, Cn)
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            '条件項目の型を指定
            'バインド変数の型を設定
            With Adapter.SelectCommand.Parameters
                .Add(New NpgsqlParameter(":YoyakuNo", NpgsqlTypes.NpgsqlDbType.Varchar))
            End With
            'バインド変数の値を設定
            With Adapter.SelectCommand
                .Parameters(":YoyakuNo").Value = YoyakuNo   '予約No
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
    ''' <param name="YoyakuNo"></param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>ログインデータを取得するSQLの作成
    ''' <para>作成情報：2015/08/04 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function setSelectRiyobiData(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByRef Cn As NpgsqlConnection, _
                                       ByVal YoyakuNo As String) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = String.Empty

        Try
            'SQL文(SELECT)
            'データアダプタに、SQLのSELECT文を設定
            strSQL = strEX04S002

            'Dim cmd = New OleDbCommand(strSQL, Cn)
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            '条件項目の型を指定
            'バインド変数の型を設定
            With Adapter.SelectCommand.Parameters
                .Add(New NpgsqlParameter(":YoyakuNo", NpgsqlTypes.NpgsqlDbType.Varchar))
            End With
            'バインド変数の値を設定
            With Adapter.SelectCommand
                .Parameters(":YoyakuNo").Value = YoyakuNo   '予約No
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
    ''' <param name="dataEXTB0102"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' <para>作成情報：2015/08/05 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function registerYoyakuInfo(ByRef Cmd As NpgsqlCommand, _
                                   ByRef Cn As NpgsqlConnection, _
                                   ByVal dataEXTB0102 As DataEXTB0102, _
                                   ByVal blnIsUpdate As Boolean) As Boolean

        Dim strSQL As String = String.Empty
        Try
            'SQL文を設定
            If blnIsUpdate = True Then
                strSQL = strEX04U001
            Else
                strSQL = strEX04I001
            End If
            Cmd = New NpgsqlCommand(strSQL, Cn)

            '値を設定
            Cmd.Parameters.Add(New NpgsqlParameter(":KariukeDt", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":KariUsercd", NpgsqlTypes.NpgsqlDbType.Varchar))
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
            'Cmd.Parameters.Add(New NpgsqlParameter(":RiyoTel22", NpgsqlTypes.NpgsqlDbType.Varchar)) '2016.11.4 m.hayabuchi del 課題No.59
            'Cmd.Parameters.Add(New NpgsqlParameter(":RiyoTel23", NpgsqlTypes.NpgsqlDbType.Varchar)) '2016.11.4 m.hayabuchi del 課題No.59
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
            'Cmd.Parameters.Add(New NpgsqlParameter(":TotalRiyoKin", NpgsqlTypes.NpgsqlDbType.Integer))                       ' 2015.12.21 UPD h.hagiwara
            Cmd.Parameters.Add(New NpgsqlParameter(":TotalRiyoKin", NpgsqlTypes.NpgsqlDbType.Bigint))                         ' 2015.12.21 UPD h.hagiwara
            Cmd.Parameters.Add(New NpgsqlParameter(":SendKbn", NpgsqlTypes.NpgsqlDbType.Char))
            Cmd.Parameters.Add(New NpgsqlParameter(":SendSts", NpgsqlTypes.NpgsqlDbType.Char))
            Cmd.Parameters.Add(New NpgsqlParameter(":SendDt", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":HensoDt", NpgsqlTypes.NpgsqlDbType.Varchar))
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
            'Cmd.Parameters(0).Value = dataEXTB0102.PropStrKariukeDt
            'Cmd.Parameters(1).Value = dataEXTB0102.PropStrKariUsercd
            'Cmd.Parameters(2).Value = DBNull.Value
            'Cmd.Parameters(3).Value = DBNull.Value
            'Cmd.Parameters(4).Value = dataEXTB0102.PropStrYoyakuSts
            'Cmd.Parameters(5).Value = dataEXTB0102.PropStrShisetuKbn
            'Cmd.Parameters(6).Value = dataEXTB0102.PropStrStudioKbn
            'Cmd.Parameters(7).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrSaijiNm)
            'Cmd.Parameters(8).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrShutsuenNm)
            'Cmd.Parameters(9).Value = dataEXTB0102.PropStrKashiKind
            'Cmd.Parameters(10).Value = dataEXTB0102.PropStrRiyoType
            'Cmd.Parameters(11).Value = dataEXTB0102.PropStrDrinkFlg
            'Cmd.Parameters(12).Value = dataEXTB0102.PropStrSaijiBunrui
            'Cmd.Parameters(13).Value = commonLogicEXT.convInsNumStr(dataEXTB0102.PropStrTeiin)
            'Cmd.Parameters(14).Value = DBNull.Value
            'Cmd.Parameters(15).Value = dataEXTB0102.PropStrRiyoshaCd
            'Cmd.Parameters(16).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrRiyoNm)
            'Cmd.Parameters(17).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrRiyoKana)
            'Cmd.Parameters(18).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrSekininBushoNm)
            'Cmd.Parameters(19).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrSekininNm)
            'Cmd.Parameters(20).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrSekininMail)
            'Cmd.Parameters(21).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrDaihyoNm)
            'Cmd.Parameters(22).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrRiyoTel11)
            'Cmd.Parameters(23).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrRiyoTel12)
            'Cmd.Parameters(24).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrRiyoTel13)
            'Cmd.Parameters(25).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrRiyoTel21)
            'Cmd.Parameters(26).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrRiyoTel22)
            'Cmd.Parameters(27).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrRiyoTel23)
            'Cmd.Parameters(28).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrRiyoNaisen)
            'Cmd.Parameters(29).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrRiyoFax11)
            'Cmd.Parameters(30).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrRiyoFax12)
            'Cmd.Parameters(31).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrRiyoFax13)
            'Cmd.Parameters(32).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrRiyoYubin1)
            'Cmd.Parameters(33).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrRiyoYubin2)
            'Cmd.Parameters(34).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrRiyoTodo)
            'Cmd.Parameters(35).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrRiyoShiku)
            'Cmd.Parameters(36).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrRiyoBan)
            'Cmd.Parameters(37).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrRiyoBuild)
            'Cmd.Parameters(38).Value = dataEXTB0102.PropStrRiyoLvl
            'Cmd.Parameters(39).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrAiteCd)
            'Cmd.Parameters(40).Value = DBNull.Value
            'Cmd.Parameters(41).Value = DBNull.Value
            'Cmd.Parameters(42).Value = DBNull.Value
            'Cmd.Parameters(43).Value = DBNull.Value
            'Cmd.Parameters(44).Value = DBNull.Value
            'Cmd.Parameters(45).Value = DBNull.Value
            'Cmd.Parameters(46).Value = DBNull.Value
            'Cmd.Parameters(47).Value = DBNull.Value
            'Cmd.Parameters(48).Value = DBNull.Value
            'Cmd.Parameters(49).Value = DBNull.Value
            'Cmd.Parameters(50).Value = DBNull.Value
            'Cmd.Parameters(51).Value = dataEXTB0102.PropStrSendKbn
            'Cmd.Parameters(52).Value = dataEXTB0102.PropStrSendSts
            'Cmd.Parameters(53).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrSendDt)
            'Cmd.Parameters(54).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrHensoDt)
            'Cmd.Parameters(55).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrBiko)
            'Cmd.Parameters(56).Value = DBNull.Value
            'Cmd.Parameters(57).Value = DBNull.Value
            'Cmd.Parameters(58).Value = DBNull.Value
            'Cmd.Parameters(59).Value = DBNull.Value
            'Cmd.Parameters(60).Value = DBNull.Value
            'Cmd.Parameters(61).Value = DBNull.Value
            'Cmd.Parameters(62).Value = DBNull.Value
            'Cmd.Parameters(63).Value = DBNull.Value
            'Cmd.Parameters(64).Value = DBNull.Value
            'Cmd.Parameters(65).Value = CommonEXT.PropComStrUserId
            'Cmd.Parameters(66).Value = CommonEXT.PropComStrUserId
            'Cmd.Parameters(67).Value = dataEXTB0102.PropStrYoyakuNo
            Cmd.Parameters(0).Value = dataEXTB0102.PropStrKariukeDt
            Cmd.Parameters(1).Value = dataEXTB0102.PropStrKariUsercd
            Cmd.Parameters(2).Value = DBNull.Value
            Cmd.Parameters(3).Value = DBNull.Value
            Cmd.Parameters(4).Value = dataEXTB0102.PropStrYoyakuSts
            Cmd.Parameters(5).Value = dataEXTB0102.PropStrShisetuKbn
            Cmd.Parameters(6).Value = dataEXTB0102.PropStrStudioKbn
            Cmd.Parameters(7).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrSaijiNm)
            Cmd.Parameters(8).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrShutsuenNm)
            Cmd.Parameters(9).Value = dataEXTB0102.PropStrKashiKind
            Cmd.Parameters(10).Value = dataEXTB0102.PropStrRiyoType
            Cmd.Parameters(11).Value = dataEXTB0102.PropStrDrinkFlg
            Cmd.Parameters(12).Value = dataEXTB0102.PropStrSaijiBunrui
            Cmd.Parameters(13).Value = commonLogicEXT.convInsNumStr(dataEXTB0102.PropStrTeiin)
            Cmd.Parameters(14).Value = DBNull.Value
            Cmd.Parameters(15).Value = dataEXTB0102.PropStrRiyoshaCd
            Cmd.Parameters(16).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrRiyoNm)
            Cmd.Parameters(17).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrRiyoKana)
            Cmd.Parameters(18).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrSekininBushoNm)
            Cmd.Parameters(19).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrSekininNm)
            Cmd.Parameters(20).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrSekininMail)
            Cmd.Parameters(21).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrDaihyoNm)
            Cmd.Parameters(22).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrRiyoTel11)
            Cmd.Parameters(23).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrRiyoTel12)
            Cmd.Parameters(24).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrRiyoTel13)
            Cmd.Parameters(25).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrRiyoTel21)
            Cmd.Parameters(26).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrRiyoNaisen)
            Cmd.Parameters(27).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrRiyoFax11)
            Cmd.Parameters(28).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrRiyoFax12)
            Cmd.Parameters(29).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrRiyoFax13)
            Cmd.Parameters(30).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrRiyoYubin1)
            Cmd.Parameters(31).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrRiyoYubin2)
            Cmd.Parameters(32).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrRiyoTodo)
            Cmd.Parameters(33).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrRiyoShiku)
            Cmd.Parameters(34).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrRiyoBan)
            Cmd.Parameters(35).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrRiyoBuild)
            Cmd.Parameters(36).Value = dataEXTB0102.PropStrRiyoLvl
            Cmd.Parameters(37).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrAiteCd)
            Cmd.Parameters(38).Value = DBNull.Value
            Cmd.Parameters(39).Value = DBNull.Value
            Cmd.Parameters(40).Value = DBNull.Value
            Cmd.Parameters(41).Value = DBNull.Value
            Cmd.Parameters(42).Value = DBNull.Value
            Cmd.Parameters(43).Value = DBNull.Value
            Cmd.Parameters(44).Value = DBNull.Value
            Cmd.Parameters(45).Value = DBNull.Value
            Cmd.Parameters(46).Value = DBNull.Value
            Cmd.Parameters(47).Value = DBNull.Value
            Cmd.Parameters(48).Value = DBNull.Value
            Cmd.Parameters(49).Value = dataEXTB0102.PropStrSendKbn
            Cmd.Parameters(50).Value = dataEXTB0102.PropStrSendSts
            Cmd.Parameters(51).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrSendDt)
            Cmd.Parameters(52).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrHensoDt)
            Cmd.Parameters(53).Value = commonLogicEXT.convInsStr(dataEXTB0102.PropStrBiko)
            Cmd.Parameters(54).Value = DBNull.Value
            Cmd.Parameters(55).Value = DBNull.Value
            Cmd.Parameters(56).Value = DBNull.Value
            Cmd.Parameters(57).Value = DBNull.Value
            Cmd.Parameters(58).Value = DBNull.Value
            Cmd.Parameters(59).Value = DBNull.Value
            Cmd.Parameters(60).Value = DBNull.Value
            Cmd.Parameters(61).Value = DBNull.Value
            Cmd.Parameters(62).Value = DBNull.Value
            Cmd.Parameters(63).Value = CommonEXT.PropComStrUserId
            Cmd.Parameters(64).Value = CommonEXT.PropComStrUserId
            Cmd.Parameters(65).Value = dataEXTB0102.PropStrYoyakuNo
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
    ''' <para>作成情報：2015/08/05 k.machida28
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function deleteYoyakuList(ByRef Cmd As NpgsqlCommand, _
                                   ByRef Cn As NpgsqlConnection, _
                                   ByVal YoyakuNo As String) As Boolean

        Dim strSQL As String = String.Empty
        Try
            'SQL文を設定
            strSQL = strEX04D002
            Cmd = New NpgsqlCommand(strSQL, Cn)

            '値を設定
            Cmd.Parameters.Add(New NpgsqlParameter(":YoyakuNo", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters(0).Value = YoyakuNo      '予約No

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
    ''' <returns></returns>
    ''' <remarks>
    ''' <para>作成情報：2015/08/05 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function insertYoyakuList(ByRef Cmd As NpgsqlCommand, _
                                   ByRef Cn As NpgsqlConnection, _
                                   ByVal dataRiyobi As CommonDataRiyobi, _
                                   ByVal YoyakuNo As String) As Boolean

        Dim strSQL As String = String.Empty
        Try
            'SQL文を設定
            strSQL = strEX04I003
            Cmd = New NpgsqlCommand(strSQL, Cn)

            '値を設定
            Cmd.Parameters.Add(New NpgsqlParameter(":ShisetuKbn", NpgsqlTypes.NpgsqlDbType.Char))
            Cmd.Parameters.Add(New NpgsqlParameter(":StudioKbn", NpgsqlTypes.NpgsqlDbType.Char))
            Cmd.Parameters.Add(New NpgsqlParameter(":YoyakuDt", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":StartTime", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":EndTime", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":YoyakuNo", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":RiyoKeitai", NpgsqlTypes.NpgsqlDbType.Char))
            Cmd.Parameters.Add(New NpgsqlParameter(":MiteiFlg", NpgsqlTypes.NpgsqlDbType.Char))
            Cmd.Parameters.Add(New NpgsqlParameter(":Tanka", NpgsqlTypes.NpgsqlDbType.Integer))
            'Cmd.Parameters.Add(New NpgsqlParameter(":Bairitu", NpgsqlTypes.NpgsqlDbType.Numeric))      ' 2015.12.02 UPD h.hagiwara
            Cmd.Parameters.Add(New NpgsqlParameter(":Bairitu", NpgsqlTypes.NpgsqlDbType.Double))        ' 2015.12.02 UPD h.hagiwara
            Cmd.Parameters.Add(New NpgsqlParameter(":Su", NpgsqlTypes.NpgsqlDbType.Integer))
            'Cmd.Parameters.Add(New NpgsqlParameter(":RiyoKin", NpgsqlTypes.NpgsqlDbType.Integer))                ' 2015.12.21 UPD h.hagiwara
            Cmd.Parameters.Add(New NpgsqlParameter(":RiyoKin", NpgsqlTypes.NpgsqlDbType.Bigint))                  ' 2015.12.21 UPD h.hagiwara
            Cmd.Parameters.Add(New NpgsqlParameter(":AddUserCd", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters.Add(New NpgsqlParameter(":UpUserCd", NpgsqlTypes.NpgsqlDbType.Varchar))

            Cmd.Parameters(0).Value = dataRiyobi.PropStrShisetuKbn
            Cmd.Parameters(1).Value = dataRiyobi.PropStrStudioKbn
            Cmd.Parameters(2).Value = dataRiyobi.PropStrYoyakuDt
            Cmd.Parameters(3).Value = dataRiyobi.PropStrStartTime
            Cmd.Parameters(4).Value = dataRiyobi.PropStrEndTime
            Cmd.Parameters(5).Value = YoyakuNo
            Cmd.Parameters(6).Value = dataRiyobi.PropStrRiyoKeitai
            Cmd.Parameters(7).Value = dataRiyobi.PropStrMiteiFlg
            Cmd.Parameters(8).Value = commonLogicEXT.convInsNum(dataRiyobi.PropIntTanka)
            'Cmd.Parameters(9).Value = commonLogicEXT.convInsNum(dataRiyobi.PropDblBairitu)       ' 2015.12.02 UPD h.hagiwara             
            Cmd.Parameters(9).Value = Double.Parse(dataRiyobi.PropDblBairitu)                     ' 2015.12.02 UPD h.hagiwara
            Cmd.Parameters(10).Value = commonLogicEXT.convInsNum(dataRiyobi.PropIntSu)
            'Cmd.Parameters(11).Value = commonLogicEXT.convInsNum(dataRiyobi.PropIntRiyoKin)       ' 2015.12.21 UPD h.hagiwara
            Cmd.Parameters(11).Value = commonLogicEXT.convLong(dataRiyobi.PropIntRiyoKin)          ' 2015.12.21 UPD h.hagiwara
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
                                   ByVal YoyakuNo As String) As Boolean

        Dim strSQL As String = String.Empty
        Try
            'SQL文を設定
            strSQL = strEX04D001
            Cmd = New NpgsqlCommand(strSQL, Cn)

            '値を設定
            Cmd.Parameters.Add(New NpgsqlParameter(":YoyakuNo", NpgsqlTypes.NpgsqlDbType.Varchar))
            Cmd.Parameters(0).Value = YoyakuNo      '予約No

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
    ''' 利用者情報取得
    ''' </summary>
    ''' <param name="Adapter"></param>
    ''' <param name="Cn"></param>
    ''' <param name="strRiyoshaKana"></param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>ログインデータを取得するSQLの作成
    ''' <para>作成情報：2015/08/04 k.machida
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function setSelectRiyosha(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByRef Cn As NpgsqlConnection, _
                                       ByVal strRiyoshaKana As String) As Boolean

        'ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        Dim strSQL As String = String.Empty

        Try
            'SQL文(SELECT)
            'データアダプタに、SQLのSELECT文を設定
            strSQL = strEX04S003
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            '条件項目の型を指定
            'バインド変数の型を設定
            With Adapter.SelectCommand.Parameters
                .Add(New NpgsqlParameter(":RiyoshaNm", NpgsqlTypes.NpgsqlDbType.Varchar))
            End With
            'バインド変数の値を設定
            With Adapter.SelectCommand
                .Parameters(":RiyoshaNm").Value = strRiyoshaKana
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
    ''' <param name="dataEXTB0102"></param>
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>責任者名から責任者メールアドレス・携帯電話番号を取得するSQL作成
    ''' <para>作成情報：2016/08/08 e.watanabe
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function GetSqlSekininshaMailTelData(ByRef Adapter As NpgsqlDataAdapter, _
                                       ByRef Cn As NpgsqlConnection, _
                                       ByVal dataEXTB0102 As DataEXTB0102) As Boolean

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
                .Parameters(":SekininNm").Value = dataEXTB0102.PropStrSekininNm   '責任者名
                .Parameters(":RiyoshaCd").Value = dataEXTB0102.PropStrRiyoshaCd   '利用者コード 
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
