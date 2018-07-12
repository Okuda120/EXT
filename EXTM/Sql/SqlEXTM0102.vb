Imports Npgsql
Imports Common


Public Class SqlEXTM0102

    Private dataEXTM0102 As DataEXTM0102

    ' 分類一覧取得ＳＱＬ
    'Private strSelectAllBunrui =
    '"SELECT " +
    '     "t1.kikan_from ||" +
    '     "t1.kikan_to                                           対象期間," +
    '     "SUBSTRING(t1.kikan_from, 1, 4)                        From_Year," +
    '     "SUBSTRING(t1.kikan_from, 6, 2)                        From_Month," +
    '     "SUBSTRING(t1.kikan_to, 1, 4)                          To_Year," +
    '     "SUBSTRING(t1.kikan_to, 6, 2)                          To_Month," +
    '     "t1.shisetu_kbn                                        施設区分, " +
    '     "t1.bunrui_cd                                          分類CD," +
    '     "t1.bunrui_nm                                          分類名," +
    '     "t1.shukei_grp                                         集計グループ," +
    '     "t1.notax_flg                                          税無フラグ," +
    '     "t1.sort                                               並び順," +
    '     "t1.kamoku_cd                                          勘定科目CD," +
    '     "t1.saimoku_cd                                         細目CD," +
    '     "t1.uchi_cd                                            内訳CD," +
    '     "t1.shosai_cd                                          詳細CD," +
    '     "t1.karikamoku_cd                                      借方勘定科目CD," +
    '     "t1.kari_saimoku_cd                                    借方細目CD," +
    '     "t1.kari_uchi_cd                                       借方内訳CD," +
    '     "t1.kari_shosai_cd                                     借方詳細CD," +
    '     "t1.sts                                                ステータス," +
    '     "t2.kamoku_nm			                            	科目名," +
    '     "t2.saimoku_nm				                            細目名," +
    '     "t2.uchi_nm				                            内訳名," +
    '     "t2.shosai_nm				                            詳細名," +
    '     "t2.karikamoku_nm				                        借方科目名,	" +
    '     "t2.kari_saimoku_nm				                    借方細目名," +
    '     "t2.kari_uchi_nm				                        借方内訳名," +
    '     "t2.kari_shosai_nm				                        借方詳細名" +
    ' " FROM " +
    '         "fbunrui_mst t1 " +
    ' "LEFT JOIN kamoku_mst t2 " +
    ' "ON t1.kamoku_cd = t2.kamoku_cd "
    ' 2015.11.30 UPD START↓ h.hagiwara 
    'Private strSelectAllBunrui =
    '"SELECT " +
    '     "t1.kikan_from ||" +
    '     "t1.kikan_to                                           対象期間," +
    '     "SUBSTRING(t1.kikan_from, 1, 4)                        From_Year," +
    '     "SUBSTRING(t1.kikan_from, 6, 2)                        From_Month," +
    '     "SUBSTRING(t1.kikan_to, 1, 4)                          To_Year," +
    '     "SUBSTRING(t1.kikan_to, 6, 2)                          To_Month," +
    '     "t1.shisetu_kbn                                        施設区分, " +
    '     "t1.bunrui_cd                                          分類CD," +
    '     "t1.bunrui_nm                                          分類名," +
    '     "t1.shukei_grp                                         集計グループ," +
    '     "t1.notax_flg                                          税無フラグ," +
    '     "t1.sort                                               並び順," +
    '     "t1.kamoku_cd                                          勘定科目CD," +
    '     "t1.saimoku_cd                                         細目CD," +
    '     "t1.uchi_cd                                            内訳CD," +
    '     "t1.shosai_cd                                          詳細CD," +
    '     "t1.karikamoku_cd                                      借方勘定科目CD," +
    '     "t1.kari_saimoku_cd                                    借方細目CD," +
    '     "t1.kari_uchi_cd                                       借方内訳CD," +
    '     "t1.kari_shosai_cd                                     借方詳細CD," +
    '     "t1.sts                                                ステータス" +
    ' " FROM " +
    '         "fbunrui_mst t1 "
    Private strSelectAllBunrui =
    "SELECT " +
         "t1.kikan_from ||" +
         "t1.kikan_to                                           対象期間," +
         "SUBSTRING(t1.kikan_from, 1, 4)                        From_Year," +
         "SUBSTRING(t1.kikan_from, 6, 2)                        From_Month," +
         "SUBSTRING(t1.kikan_to, 1, 4)                          To_Year," +
         "SUBSTRING(t1.kikan_to, 6, 2)                          To_Month," +
         "t1.shisetu_kbn                                        施設区分, " +
         "t1.bunrui_cd                                          分類CD," +
         "t1.bunrui_nm                                          分類名," +
         "t1.shukei_grp                                         集計グループ," +
         "t1.notax_flg                                          税無フラグ," +
         "t1.sort                                               並び順," +
         "t1.kamoku_cd                                          勘定科目CD," +
         "t1.saimoku_cd                                         細目CD," +
         "t1.uchi_cd                                            内訳CD," +
         "t1.shosai_cd                                          詳細CD," +
         "t1.karikamoku_cd                                      借方勘定科目CD," +
         "t1.kari_saimoku_cd                                    借方細目CD," +
         "t1.kari_uchi_cd                                       借方内訳CD," +
         "t1.kari_shosai_cd                                     借方詳細CD," +
         "t1.sts                                                ステータス," +
         "t1.RIYORYO_FLG                                        付帯利用料フラグ" +
     " FROM " +
             "fbunrui_mst t1 "
    ' 2015.11.30 UPD END↑ h.hagiwara 

    '付帯設備の取得ＳＱＬ
    Private strSelectFutai As String =
                                        "SELECT " +
                                            " futai_cd," +
                                            " futai_nm," +
                                            " tanka," +
                                            " tani," +
                                            " def_flg," +
                                            " sort," +
                                            " sts," +
                                            " 0," +
                                            " bunrui_cd," +
                                            "'0'" +
                                       " FROM" +
                                            " futai_mst"


    '対象期間取得SQL
    Private strSelectFinishedFromTo As String =
                                                    "SELECT " +
                                                    " DISTINCT " +
                                                    " kikan_from || kikan_to  対象期間コード," +
                                                    "kikan_from || '-'" +
                                                     "  || kikan_to           対象期間" +
                                                    " FROM " +
                                                     " fbunrui_mst "


    '勘定科目の取得ＳＱＬ
    ' 2015.10.09 UPDATE START↓ h.hagiwara 勘定科目の重複取得対応
    'Private strSelectKanzyoKmk As String =
    '                                        "SELECT " +
    '                                            "kamoku_cd," +
    '                                            "kamoku_nm" +
    '                                        " FROM" +
    '                                            " kamoku_mst"
    ' 2016.06.24 UPD START↓ h.hagiwara コンボリスト設定方法変更対応
    'Private strSelectKanzyoKmk As String =
    '                                        "SELECT DISTINCT " +
    '                                            "kamoku_cd," +
    '                                            "kamoku_nm" +
    '                                        " FROM" +
    '                                            " kamoku_mst"
    ' 2015.10.09 UPDATE END↑ h.hagiwara 勘定科目の重複取得対応
    Private strSelectKanzyoKmk As String =
                                            "SELECT DISTINCT" & vbCrLf & _
                                            "    KAMOKU_CD," & vbCrLf & _
                                            "    KAMOKU_NM" & vbCrLf & _
                                            " FROM" & vbCrLf & _
                                            "    kamoku_mst" & vbCrLf & _
                                            " WHERE " & vbCrLf & _
                                            "    sts = '0'" & vbCrLf & _
                                            " ORDER BY " & vbCrLf & _
                                            "    KAMOKU_CD "
    ' 2016.06.24 UPD END↑ h.hagiwara コンボリスト設定方法変更対応

    ''細目の取得
    '' 2015.10.09 UPDATE START↓ h.hagiwara 勘定科目の重複取得対応
    ''Private strSelectSaimoku As String =
    ''                                        "SELECT " +
    ''                                            " saimoku_cd," +
    ''                                            " saimoku_nm" +
    ''                                        " FROM" +
    ''                                            " kamoku_mst"
    ' 2016.06.24 UPD START↓ h.hagiwara コンボリスト設定方法変更対応
    'Private strSelectSaimoku As String =
    '                                        "SELECT DISTINCT " +
    '                                            " saimoku_cd," +
    '                                            " saimoku_nm" +
    '                                        " FROM" +
    '                                            " kamoku_mst"
    '' 2015.10.09 UPDATE END↑ h.hagiwara 勘定科目の重複取得対応
    Private strSelectSaimoku As String =
                                            "SELECT DISTINCT " & vbCrLf & _
                                            "    KAMOKU_CD," & vbCrLf & _
                                            "    SAIMOKU_CD," & vbCrLf & _
                                            "    SAIMOKU_NM" & vbCrLf & _
                                            " FROM" & vbCrLf & _
                                            "    kamoku_mst" & vbCrLf & _
                                            " WHERE " & vbCrLf & _
                                            "    sts = '0'" & vbCrLf & _
                                            " ORDER BY " & vbCrLf & _
                                            "    KAMOKU_CD ," & vbCrLf & _
                                            "    SAIMOKU_CD "
    ' 2016.06.24 UPD END↑ h.hagiwara コンボリスト設定方法変更対応

    ''内訳の取得
    '' 2015.10.09 UPDATE START↓ h.hagiwara 勘定科目の重複取得対応
    ''Private strSelectUchiwake As String =
    ''                                        "SELECT " +
    ''                                            "uchi_cd," +
    ''                                            "uchi_nm" +
    ''                                        " FROM" +
    ''                                            " kamoku_mst"
    ' 2016.06.24 UPD START↓ h.hagiwara コンボリスト設定方法変更対応
    'Private strSelectUchiwake As String =
    '                                        "SELECT DISTINCT " +
    '                                            "uchi_cd," +
    '                                            "uchi_nm" +
    '                                        " FROM" +
    '                                            " kamoku_mst"
    '' 2015.10.09 UPDATE END↑ h.hagiwara 勘定科目の重複取得対応
    Private strSelectUchiwake As String =
                                            "SELECT DISTINCT " & vbCrLf & _
                                            "    KAMOKU_CD," & vbCrLf & _
                                            "    SAIMOKU_CD," & vbCrLf & _
                                            "    UCHI_CD," & vbCrLf & _
                                            "    UCHI_NM" & vbCrLf & _
                                            " FROM" & vbCrLf & _
                                            "    kamoku_mst" & vbCrLf & _
                                            " WHERE " & vbCrLf & _
                                            "    sts = '0'" & vbCrLf & _
                                            " ORDER BY " & vbCrLf & _
                                            "    KAMOKU_CD ," & vbCrLf & _
                                            "    SAIMOKU_CD ," & vbCrLf & _
                                            "    UCHI_CD "
    ' 2016.06.24 UPD END↑ h.hagiwara コンボリスト設定方法変更対応


    ''詳細の取得
    ' 2016.06.24 UPD START↓ h.hagiwara コンボリスト設定方法変更対応
    'Private strSelectShosai As String =
    '                                        "SELECT " +
    '                                            "shosai_cd," +
    '                                            "shosai_nm" +
    '                                        " FROM" +
    '                                            " kamoku_mst"
    ' 2016.06.24 DELETE END↑ h.hagiwara コンボリスト設定方法変更対応
    Private strSelectShosai As String =
                                            "SELECT " & vbCrLf & _
                                            "    KAMOKU_CD," & vbCrLf & _
                                            "    SAIMOKU_CD," & vbCrLf & _
                                            "    UCHI_CD," & vbCrLf & _
                                            "    SHOSAI_CD," & vbCrLf & _
                                            "    SHOSAI_NM" & vbCrLf & _
                                            " FROM" & vbCrLf & _
                                            "    kamoku_mst" & vbCrLf & _
                                            " WHERE " & vbCrLf & _
                                            "    sts = '0'" & vbCrLf & _
                                            " ORDER BY " & vbCrLf & _
                                            "    KAMOKU_CD ," & vbCrLf & _
                                            "    SAIMOKU_CD ," & vbCrLf & _
                                            "    UCHI_CD ," & vbCrLf & _
                                            "    SHOSAI_CD"
    ' 2016.06.24 UPD END↑ h.hagiwara コンボリスト設定方法変更対応


    ' 2016.04.28 DEL START↓ h.hagiwara レスポンス改善
    ''借方勘定科目の取得
    '' 2015.10.09 UPDATE START↓ h.hagiwara 勘定科目の重複取得対応
    ''Private strSelectKariKanzyo As String =
    ''                                        "SELECT " +
    ''                                            "karikamoku_cd," +
    ''                                            "karikamoku_nm" +
    ''                                        " FROM" +
    ''                                            " kamoku_mst"
    'Private strSelectKariKanzyo As String =
    '                                        "SELECT DISTINCT " +
    '                                            "karikamoku_cd," +
    '                                            "karikamoku_nm" +
    '                                        " FROM" +
    '                                            " kamoku_mst"
    '' 2015.10.09 UPDATE END↑ h.hagiwara 勘定科目の重複取得対応

    ''借方細目の取得
    '' 2015.10.09 UPDATE START↓ h.hagiwara 勘定科目の重複取得対応
    ''Private strSelectKariSaimoku As String =
    ''                                        "SELECT " +
    ''                                            "kari_saimoku_cd," +
    ''                                            "kari_saimoku_nm" +
    ''                                        " FROM" +
    ''                                            " kamoku_mst"
    'Private strSelectKariSaimoku As String =
    '                                        "SELECT DISTINCT " +
    '                                            "kari_saimoku_cd," +
    '                                            "kari_saimoku_nm" +
    '                                        " FROM" +
    '                                            " kamoku_mst"
    '' 2015.10.09 UPDATE END↑ h.hagiwara 勘定科目の重複取得対応

    ''借方内訳の取得
    '' 2015.10.09 UPDATE START↓ h.hagiwara 勘定科目の重複取得対応
    ''Private strSelectKariUchi As String =
    ''                                        "SELECT " +
    ''                                            "kari_uchi_cd," +
    ''                                            "kari_uchi_nm" +
    ''                                        " FROM" +
    ''                                            " kamoku_mst"
    'Private strSelectKariUchi As String =
    '                                        "SELECT DISTINCT " +
    '                                            "kari_uchi_cd," +
    '                                            "kari_uchi_nm" +
    '                                        " FROM" +
    '                                            " kamoku_mst"
    '' 2015.10.09 DEL END↑ h.hagiwara 勘定科目の重複取得対応

    ''借方詳細の取得
    'Private strSelectKariShosai As String =
    '                                        "SELECT " +
    '                                            "kari_shosai_cd," +
    '                                            "kari_shosai_nm" +
    '                                        " FROM" +
    '                                            " kamoku_mst"
    ' 2016.04.28 DEL END↑ h.hagiwara レスポンス改善

    ' 2015.11.30 UPD START↓ h.hagiwara
    '分類登録sql
    'Private strInsertBunrui As String =
    '                                        "INSERT INTO " +
    '                                        " fbunrui_mst" +
    '                                        "( kikan_from," +
    '                                        "  kikan_to," +
    '                                        "  shisetu_kbn," +
    '                                        "  bunrui_cd," +
    '                                        "  bunrui_nm," +
    '                                        "  shukei_grp," +
    '                                        "  notax_flg," +
    '                                        "  sort," +
    '                                        "  kamoku_cd," +
    '                                        "  saimoku_cd," +
    '                                        "  uchi_cd," +
    '                                        "  shosai_cd," +
    '                                        "  karikamoku_cd," +
    '                                        "  kari_saimoku_cd," +
    '                                        "  kari_uchi_cd," +
    '                                        "  kari_shosai_cd," +
    '                                        "  sts," +
    '                                        "  add_dt," +
    '                                        "  add_user_cd," +
    '                                        "  up_dt," +
    '                                        "  up_user_cd" +
    '                                        " )VALUES( " +
    '                                        "  :kikanFrom," +
    '                                        "  :kikanTo," +
    '                                        "  :shisetuKbn," +
    '                                        "  :bunruiCd," +
    '                                        "  :bunruiNm," +
    '                                        "  :shukeiGrp," +
    '                                        "  :notaxFlg," +
    '                                        "  :sort," +
    '                                        "  :kamokuCd," +
    '                                        "  :saimokuCd," +
    '                                        "  :uchiCd," +
    '                                        "  :shosaiCd," +
    '                                        "  :karikamokuCd," +
    '                                        "  :karisaimokuCd," +
    '                                        "  :kariuchiCd," +
    '                                        "  :karishosaiCd," +
    '                                        "  :sts," +
    '                                        "  :addDate," +
    '                                        "  :addUserCd," +
    '                                        "  :upDate," +
    '                                        "  :upUserCd)"
    Private strInsertBunrui As String =
                                            "INSERT INTO " +
                                            " fbunrui_mst" +
                                            "( kikan_from," +
                                            "  kikan_to," +
                                            "  shisetu_kbn," +
                                            "  bunrui_cd," +
                                            "  bunrui_nm," +
                                            "  shukei_grp," +
                                            "  notax_flg," +
                                            "  sort," +
                                            "  kamoku_cd," +
                                            "  saimoku_cd," +
                                            "  uchi_cd," +
                                            "  shosai_cd," +
                                            "  karikamoku_cd," +
                                            "  kari_saimoku_cd," +
                                            "  kari_uchi_cd," +
                                            "  kari_shosai_cd," +
                                            "  sts," +
                                            "  RIYORYO_FLG," +
                                            "  add_dt," +
                                            "  add_user_cd," +
                                            "  up_dt," +
                                            "  up_user_cd" +
                                            " )VALUES( " +
                                            "  :kikanFrom," +
                                            "  :kikanTo," +
                                            "  :shisetuKbn," +
                                            "  :bunruiCd," +
                                            "  :bunruiNm," +
                                            "  :shukeiGrp," +
                                            "  :notaxFlg," +
                                            "  :sort," +
                                            "  :kamokuCd," +
                                            "  :saimokuCd," +
                                            "  :uchiCd," +
                                            "  :shosaiCd," +
                                            "  :karikamokuCd," +
                                            "  :karisaimokuCd," +
                                            "  :kariuchiCd," +
                                            "  :karishosaiCd," +
                                            "  :sts," +
                                            "  :RIYORYO_FLG," +
                                            "  :addDate," +
                                            "  :addUserCd," +
                                            "  :upDate," +
                                            "  :upUserCd)"
    ' 2015.11.30 UPD END↑ h.hagiwara

    ' 2015.11.30 UPD START↓ h.hagiwara
    '分類更新sql
    'Private strUpdateBunrui As String =
    '                                        "UPDATE " +
    '                                        " fbunrui_mst " +
    '                                        " SET " +
    '                                        "  kikan_from = :kikanFrom," +
    '                                        "  kikan_to   = :kikanTo," +
    '                                        "  bunrui_nm = :bunruiNm," +
    '                                        "  shukei_grp = :shukeiGrp," +
    '                                        "  notax_flg = :notaxFlg," +
    '                                        "  sort = :sort," +
    '                                        "  kamoku_cd = :kamokuCd," +
    '                                        "  saimoku_cd = :saimokuCd," +
    '                                        "  uchi_cd = :uchiCd," +
    '                                        "  shosai_cd = :shosaiCd," +
    '                                        "  karikamoku_cd = :karikamokuCd," +
    '                                        "  kari_saimoku_cd = :karisaimokuCd," +
    '                                        "  kari_uchi_cd = :kariuchiCd," +
    '                                        "  kari_shosai_cd = :karishosaiCd," +
    '                                        "  sts = :sts," +
    '                                        "  up_dt = :upDate," +
    '                                        "  up_user_cd = :upUserCd "
    Private strUpdateBunrui As String =
                                           "UPDATE " +
                                           " fbunrui_mst " +
                                           " SET " +
                                           "  kikan_from = :kikanFrom," +
                                           "  kikan_to   = :kikanTo," +
                                           "  bunrui_nm = :bunruiNm," +
                                           "  shukei_grp = :shukeiGrp," +
                                           "  notax_flg = :notaxFlg," +
                                           "  sort = :sort," +
                                           "  kamoku_cd = :kamokuCd," +
                                           "  saimoku_cd = :saimokuCd," +
                                           "  uchi_cd = :uchiCd," +
                                           "  shosai_cd = :shosaiCd," +
                                           "  karikamoku_cd = :karikamokuCd," +
                                           "  kari_saimoku_cd = :karisaimokuCd," +
                                           "  kari_uchi_cd = :kariuchiCd," +
                                           "  kari_shosai_cd = :karishosaiCd," +
                                           "  sts = :sts," +
                                           "  RIYORYO_FLG = :RIYORYO_FLG," +
                                           "  up_dt = :upDate," +
                                           "  up_user_cd = :upUserCd "
    ' 2015.11.30 UPD END↑ h.hagiwara

    '付帯設備登録sql
    Private strInsertFutai As String =
                                            "INSERT INTO " +
                                                " futai_mst( " +
                                                " kikan_from," +
                                                " kikan_to," +
                                                " shisetu_kbn," +
                                                " bunrui_cd," +
                                                " futai_cd," +
                                                " futai_nm," +
                                                " tanka," +
                                                " tani," +
                                                " sts," +
                                                " def_flg," +
                                                " sort," +
                                                " add_Dt," +
                                                " add_user_cd," +
                                                " up_dt," +
                                                " up_user_cd) " +
                                            "VALUES( " +
                                                " :kikanFrom," +
                                                " :kikanTo," +
                                                " :shisetuKbn," +
                                                " :bunruiCd ," +
                                                " :futaiCd," +
                                                " :futaiNm," +
                                                " :tanka," +
                                                " :tani," +
                                                " :sts," +
                                                " :defFlg," +
                                                " :sort," +
                                                " :addDate," +
                                                " :addUserCd," +
                                                " :upDate," +
                                                " :upUserCd)"

    '付帯設備表更新
    Private strUpdateFutai =
                                            "UPDATE " +
                                            " futai_mst " +
                                            " SET " +
                                            " kikan_from = :kikanFrom," +
                                            " kikan_to   = :kikanTo," +
                                            " futai_nm = :futaiNm," +
                                            " tanka = :tanka," +
                                            " tani = :tani," +
                                            " sts = :sts," +
                                            " def_flg = :defFlg," +
                                            " sort = :sort," +
                                            " up_dt = :upDate," +
                                            " up_user_cd = :upUserCd"

    Private strSelectbunricd =
                                            "SELECT " +
                                            " MAX(bunrui_cd) " +
                                            "FROM fbunrui_mst "
    Private strSelectfutaibunricd =
                                            "SELECT " +
                                            " MAX(futai_cd) " +
                                            "FROM futai_mst "
    ' 2015.12.25 ADD START↓ h.hagiwara
    ' 利用料用の情報が存在するか科目マスタから取得
    Private strSelectKamoku = " SELECT RIYO_KAMOKU_FLG " & vbCrLf & _
                              " FROM KAMOKU_MST " & vbCrLf & _
                              " WHERE KAMOKU_CD  = :KAMOKU_CD " & vbCrLf & _
                              "  AND  SAIMOKU_CD = :SAIMOKU_CD " & vbCrLf & _
                              "  AND  UCHI_CD    = :UCHI_CD " & vbCrLf & _
                              "  AND  SHOSAI_CD  = :SHOSAI_CD "
    ' 登録時の科目コードなどを取得
    Private strGetKamoku = " SELECT " & vbCrLf & _
                           "      * " & vbCrLf & _
                           " FROM KAMOKU_MST " & vbCrLf & _
                           " WHERE RIYO_KAMOKU_FLG  = '1' " & vbCrLf &
                           "  AND  STS = '0' "
    ' 2015.12.25 ADD END↑ h.hagiwara

    ''' <summary>
    ''' 初期表示用SQL作成
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataHBKA0101">[IN]データクラス</param>
    ''' </summary> 
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>付帯設備分類マスタから、情報を取得し、値をセットする（今はセットしていない）
    ''' <para>作成情報：2015/08/11 yu.satoh 
    ''' </para></remarks>
    Public Function SetSelectIniData(ByRef Adapter As NpgsqlDataAdapter, ByVal Cn As NpgsqlConnection, ByVal dataEXTM0102 As DataEXTM0102) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数の宣言
        Dim strSQL As String
        Dim strWhere As String
        Dim strOrder As String
        Dim dtDate As String = Now.ToString("yyyy/MM/dd")

        Try

            'SQL文(SELECT)
            strSQL = strSelectAllBunrui
            'Where句作成
            '施設区分によって条件変更
            If dataEXTM0102.PropTheaterBtn.Checked = True Then
                strWhere = " WHERE t1.shisetu_kbn = '1' "
            Else
                strWhere = " WHERE t1.shisetu_kbn = '2' "
            End If
            '初期表示なら現在の日付を、そうでないなら対象期間を設定
            If dataEXTM0102.PropInitFlg = False Then
                strWhere &= " AND '" + dtDate + "' BETWEEN t1.kikan_from AND t1.kikan_to"
            Else
                strWhere &= " AND :item = t1.kikan_from || t1.kikan_to"
            End If

            'ORDER BY句作成
            strOrder = " ORDER BY t1.sort asc"

            'データアダプタに、SQLを設定
            strSQL &= strWhere
            strSQL &= strOrder
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            'バインド変数をセット
            Adapter.SelectCommand.Parameters.Add(New NpgsqlParameter("item", NpgsqlTypes.NpgsqlDbType.Varchar)) '対象期間
            Adapter.SelectCommand.Parameters("item").Value = dataEXTM0102.PropFinishedFromTo.SelectedValue      '対象期間
            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
            puErrMsg = EXTM0102_E0000 & ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' 初期表示用SQL作成
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataHBKA0101">[IN]データクラス</param>
    ''' </summary> 
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>付帯設備マスタから、情報を取得
    ''' <para>作成情報：2015/08/11 yu.satoh 
    ''' </para></remarks>
    Public Function SetSelectIniFutaiData(ByRef Adapter As NpgsqlDataAdapter, ByVal Cn As NpgsqlConnection, ByVal dataEXTM0102 As DataEXTM0102, ByRef row As Integer) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数の宣言
        Dim strSQL As String
        Dim strWhere As String
        Dim strOrder As String
        Dim dtDate As String = Now.ToString("yyyy/MM/dd")   '現在の日付

        Try

            'SQL文(SELECT)
            strSQL = strSelectFutai
            'Where句作成
            'シアターにチェックがあれば1、なければ2を設定
            If dataEXTM0102.PropTheaterBtn.Checked = True Then
                strWhere = " WHERE shisetu_kbn = '1' "
            Else
                strWhere = " WHERE shisetu_kbn = '2' "
            End If

            'strWhere &= " AND bunrui_cd = :bunruiCd "
            '初期表示ならば現在日付を条件に、そうでなければ対象期間を条件にする
            If dataEXTM0102.PropInitFlg = False Then
                strWhere &= " AND '" + dtDate + "' BETWEEN kikan_from AND kikan_to"
            Else
                strWhere &= " AND :item = kikan_from || kikan_to"
            End If

            'ORDER BY句作成
            strOrder = " ORDER BY bunrui_cd, sort"

            'データアダプタに、SQLを設定
            strSQL &= strWhere
            strSQL &= strOrder
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            'バインド変数をセット
            'Adapter.SelectCommand.Parameters.Add(New NpgsqlParameter("bunruiCd", NpgsqlTypes.NpgsqlDbType.Varchar)) '分類コード
            Adapter.SelectCommand.Parameters.Add(New NpgsqlParameter("item", NpgsqlTypes.NpgsqlDbType.Varchar))     '対象期間

            '初期表示では分類表の(0,0)を、そうでないなら取得したクリックした行と同じ行にある分類コードを設定
            'If row <> 0 Then
            '    Adapter.SelectCommand.Parameters("bunruiCd").Value = dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(row, 0).Text
            'Else
            '    'もし分類表に何もない場合、分類コードは-1を設定し、何も表示させない
            '    If dataEXTM0102.PropDtFbunruiMst.Rows.Count = Nothing Then
            '        Adapter.SelectCommand.Parameters("bunruiCd").Value = "-1"
            '    Else
            '        Adapter.SelectCommand.Parameters("bunruiCd").Value = dataEXTM0102.PropDtFbunruiMst.Rows(row).Item(6)
            '    End If
            'End If

            Adapter.SelectCommand.Parameters("item").Value = dataEXTM0102.PropFinishedFromTo.SelectedValue          '対象期間

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
            puErrMsg = EXTM0102_E0000 & ex.Message
            Return False
        End Try

    End Function


    ''' <summary>
    ''' 対象期間コンボボックスSQL作成
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataHBKA0101">[IN]データクラス</param>
    ''' </summary> 
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>付帯設備マスタから対象期間を取得
    ''' <para>作成情報：2015/08/14 yu.satoh 
    ''' </para></remarks>
    Public Function SetFinishedFromToData(ByRef Adapter As NpgsqlDataAdapter, ByVal Cn As NpgsqlConnection, ByVal dataEXTM0102 As DataEXTM0102) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数の宣言
        Dim strSQL As String
        Dim strWhere As String
        Dim strORDER As String

        Try

            'SQL文(SELECT)
            strSQL = strSelectFinishedFromTo
            'where句作成
            If dataEXTM0102.PropTheaterBtn.Checked = True Then
                strWhere = " WHERE shisetu_kbn = '1' "
            Else
                strWhere = " WHERE shisetu_kbn = '2' "
            End If

            'ORDER BY句作成
            strORDER = " ORDER BY 対象期間コード"

            'データアダプタに、SQLを設定
            strSQL &= strWhere
            strSQL &= strORDER
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
            puErrMsg = EXTM0102_E0000 & ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' 勘定科目コンボボックス作成SQL作成
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataHBKA0101">[IN]データクラス</param>
    ''' </summary> 
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>勘定科目マスタから勘定科目のコンボボックスをセットする
    ''' <para>作成情報：2015/08/11 yu.satoh 
    ''' </para></remarks>
    Public Function CmbKanzyoCdSet(ByRef Adapter As NpgsqlDataAdapter, ByVal Cn As NpgsqlConnection, ByVal dataEXTM0102 As DataEXTM0102)

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数の宣言
        Dim strSQL As String
        'Dim strORDER As String                                       ' 2016.06.24 DEL h.hagiwara コンボリスト設定方法変更対応

        Try

            'SQL文(SELECT)
            strSQL = strSelectKanzyoKmk
            ' 2016.06.24 DEL START↓ h.hagiwara コンボリスト設定方法変更対応
            ''orderBy句作成
            'strORDER = " ORDER BY kamoku_cd"
            ' 2016.06.24 DEL END↑ h.hagiwara コンボリスト設定方法変更対応

            'データアダプタに、SQLを設定
            'strSQL &= strORDER                                       ' 2016.06.24 DEL h.hagiwara コンボリスト設定方法変更対応
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
            puErrMsg = EXTM0102_E0000 & ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' 細目コンボボックス作成SQL作成
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataHBKA0101">[IN]データクラス</param>
    ''' </summary> 
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>勘定科目マスタから細目のコンボボックスをセットする
    ''' <para>作成情報：2015/08/11 yu.satoh 
    ''' </para></remarks>
    Public Function CmbSaimouCdSet(ByRef Adapter As NpgsqlDataAdapter, ByVal Cn As NpgsqlConnection, ByVal dataEXTM0102 As DataEXTM0102)                                  ' 2016.06.24 UPD h.hagiwara コンボリスト設定方法変更対応
        'Public Function CmbSaimouCdSet(ByRef Adapter As NpgsqlDataAdapter, ByVal Cn As NpgsqlConnection, ByVal dataEXTM0102 As DataEXTM0102, ByVal kamokuCd As String)   ' 2016.06.24 UPD h.hagiwara コンボリスト設定方法変更対応

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数の宣言
        Dim strSQL As String
        ' 2016.06.24 DEL START↓ h.hagiwara コンボリスト設定方法変更対応
        'Dim strWhere As String
        'Dim strORDER As String
        ' 2016.06.24 DEL END↑ h.hagiwara コンボリスト設定方法変更対応

        Try

            'SQL文(SELECT)
            strSQL = strSelectSaimoku
            ' 2016.06.24 DEL START↓ h.hagiwara コンボリスト設定方法変更対応
            ''Where句作成
            'strWhere = " WHERE kamoku_cd = :kamokuCd"
            ''ORDER BY句
            ''strORDER = " ORDER BY kamoku_cd"
            'strORDER = " ORDER BY saimoku_cd"
            ' 2016.06.24 DEL END↑ h.hagiwara コンボリスト設定方法変更対応

            'データアダプタに、SQLを設定
            ' 2016.06.24 DEL START↓ h.hagiwara コンボリスト設定方法変更対応
            'strSQL &= strWhere
            'strSQL &= strORDER
            ' 2016.06.24 DEL END↑ h.hagiwara コンボリスト設定方法変更対応
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            ' 2016.06.24 DEL START↓ h.hagiwara コンボリスト設定方法変更対応
            'バインド変数に型と値をセット
            'Adapter.SelectCommand.Parameters.Add(New NpgsqlParameter("kamokuCd", NpgsqlTypes.NpgsqlDbType.Varchar)) '科目コード
            'Adapter.SelectCommand.Parameters("kamokuCd").Value = kamokuCd                                           '科目コード
            ' 2016.06.24 DEL END↑ h.hagiwara コンボリスト設定方法変更対応

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
            puErrMsg = EXTM0102_E0000 & ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' 内訳コンボボックス作成SQL作成
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataHBKA0101">[IN]データクラス</param>
    ''' </summary> 
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>勘定科目マスタから内訳のコンボボックスをセットする
    ''' <para>作成情報：2015/08/11 yu.satoh 
    ''' </para></remarks>
    Public Function CmbUchiwakeCdSet(ByRef Adapter As NpgsqlDataAdapter, ByVal Cn As NpgsqlConnection, ByVal dataEXTM0102 As DataEXTM0102)                            ' 2016.06.24 UPD h.hagiwara コンボリスト設定方法変更対応
        'Public Function CmbUchiwakeCdSet(ByRef Adapter As NpgsqlDataAdapter, ByVal Cn As NpgsqlConnection, ByVal dataEXTM0102 As DataEXTM0102, ByVal kamokuCd As String, ByVal saimokuCd As String)' 2016.06.24 UPD h.hagiwara コンボリスト設定方法変更対応

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数の宣言
        Dim strSQL As String
        'Dim strWhere As String                   ' 2016.06.24 DEL h.hagiwara コンボリスト設定方法変更対応
        'Dim strORDER As String                   ' 2016.06.24 DEL h.hagiwara コンボリスト設定方法変更対応

        Try

            'SQL文(SELECT)
            strSQL = strSelectUchiwake
            ' 2016.06.24 DEL START↓ h.hagiwara コンボリスト設定方法変更対応
            ''Where句作成
            'strWhere = " WHERE kamoku_cd = :kamokuCd" +
            '           " AND saimoku_cd = :saimokuCd"
            ''ORDER BY句作成
            ''strORDER = " ORDER BY kamoku_cd"
            'strORDER = " ORDER BY uchi_cd"
            ' 2016.06.24 DEL END↑ h.hagiwara コンボリスト設定方法変更対応


            'データアダプタに、SQLを設定
            ' 2016.06.24 DEL START↓ h.hagiwara コンボリスト設定方法変更対応
            'strSQL &= strWhere
            'strSQL &= strORDER
            ' 2016.06.24 DEL END↑ h.hagiwara コンボリスト設定方法変更対応
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            ' 2016.06.24 DEL START↓ h.hagiwara コンボリスト設定方法変更対応
            'バインド変数に型と値をセット
            'Adapter.SelectCommand.Parameters.Add(New NpgsqlParameter("kamokuCd", NpgsqlTypes.NpgsqlDbType.Varchar)) '科目コード
            'Adapter.SelectCommand.Parameters("kamokuCd").Value = kamokuCd                                           '科目コード
            'Adapter.SelectCommand.Parameters.Add(New NpgsqlParameter("saimokuCd", NpgsqlTypes.NpgsqlDbType.Varchar)) '細目コード
            'Adapter.SelectCommand.Parameters("saimokuCd").Value = saimokuCd                                          '細目コード
            ' 2016.06.24 DEL END↑ h.hagiwara コンボリスト設定方法変更対応

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
            puErrMsg = EXTM0102_E0000 & ex.Message
            Return False
        End Try

    End Function


    ''' <summary>
    ''' 詳細コンボボックス作成SQL作成
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataHBKA0101">[IN]データクラス</param>
    ''' </summary> 
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>勘定科目マスタから内詳細のコンボボックスをセットする
    ''' <para>作成情報：2015/08/11 yu.satoh 
    ''' </para></remarks> 
    Public Function CmbShosaiCdSet(ByRef Adapter As NpgsqlDataAdapter, ByVal Cn As NpgsqlConnection, ByVal dataEXTM0102 As DataEXTM0102)                             ' 2016.06.24 UPD h.hagiwara コンボリスト設定方法変更対応
        'Public Function CmbShosaiCdSet(ByRef Adapter As NpgsqlDataAdapter, ByVal Cn As NpgsqlConnection, ByVal dataEXTM0102 As DataEXTM0102, ByVal kamokuCd As String, ByVal saimokuCd As String, ByVal uchiCd As String)' 2016.06.24 UPD h.hagiwara コンボリスト設定方法変更対応

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数の宣言
        Dim strSQL As String
        'Dim strWhere As String             ' 2016.06.24 DEL h.hagiwara コンボリスト設定方法変更対応
        'Dim strORDER As String             ' 2016.06.24 DEL h.hagiwara コンボリスト設定方法変更対応

        Try

            'SQL文(SELECT)
            strSQL = strSelectShosai
            ' 2016.06.24 DEL START↓ h.hagiwara コンボリスト設定方法変更対応
            ''Where句作成
            'strWhere = " WHERE kamoku_cd = :kamokuCd" +
            '           " AND saimoku_cd = :saimokuCd" +
            '           " AND uchi_cd = :uchiCd"
            ''ORDER BY句作成
            'strORDER = " ORDER BY kamoku_cd"
            ' 2016.06.24 DEL END↑ h.hagiwara コンボリスト設定方法変更対応


            'データアダプタに、SQLを設定
            ' 2016.06.24 DEL START↓ h.hagiwara コンボリスト設定方法変更対応
            'strSQL &= strWhere
            'strSQL &= strORDER
            ' 2016.06.24 DEL END↑ h.hagiwara コンボリスト設定方法変更対応
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            ' 2016.06.24 DEL START↓ h.hagiwara コンボリスト設定方法変更対応
            ''バインド変数に型と値をセット
            'Adapter.SelectCommand.Parameters.Add(New NpgsqlParameter("kamokuCd", NpgsqlTypes.NpgsqlDbType.Varchar)) '科目コード
            'Adapter.SelectCommand.Parameters("kamokuCd").Value = kamokuCd                                           '科目コード
            'Adapter.SelectCommand.Parameters.Add(New NpgsqlParameter("saimokuCd", NpgsqlTypes.NpgsqlDbType.Varchar)) '細目コード
            'Adapter.SelectCommand.Parameters("saimokuCd").Value = saimokuCd                                          '細目コード
            'Adapter.SelectCommand.Parameters.Add(New NpgsqlParameter("uchiCd", NpgsqlTypes.NpgsqlDbType.Varchar))    '内訳コード
            'Adapter.SelectCommand.Parameters("uchiCd").Value = uchiCd                                                '内訳コード
            ' 2016.06.24 DEL END↑ h.hagiwara コンボリスト設定方法変更対応

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
            puErrMsg = EXTM0102_E0000 & ex.Message
            Return False
        End Try

    End Function


    ' 2016.04.28 DEL START↓ h.hagiwara レスポンス改善
    ' ''' <summary>
    ' ''' 借方科目コンボボックス作成SQL作成
    ' ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ' ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ' ''' <param name="dataHBKA0101">[IN]データクラス</param>
    ' ''' </summary> 
    ' ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ' ''' <remarks>勘定科目マスタから借方科目のコンボボックスをセットする
    ' ''' <para>作成情報：2015/08/11 yu.satoh 
    ' ''' </para></remarks>
    'Public Function CmbKariKamokuSet(ByRef Adapter As NpgsqlDataAdapter, ByVal Cn As NpgsqlConnection, ByVal dataEXTM0102 As DataEXTM0102)

    '    '開始ログ出力
    '    CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

    '    '変数の宣言
    '    Dim strSQL As String
    '    Dim strORDER As String

    '    Try

    '        'SQL文(SELECT)
    '        strSQL = strSelectKariKanzyo
    '        'ORDER BY句作成
    '        strORDER = " ORDER BY karikamoku_cd"

    '        'データアダプタに、SQLを設定
    '        strSQL &= strORDER
    '        Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

    '        '終了ログ出力
    '        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

    '        '正常終了
    '        Return True

    '    Catch ex As Exception
    '        '例外発生
    '        CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
    '        puErrMsg = EXTM0102_E0000 & ex.Message
    '        Return False
    '    End Try

    'End Function

    ' ''' <summary>
    ' ''' 借方細目コンボボックス作成SQL作成
    ' ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ' ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ' ''' <param name="dataHBKA0101">[IN]データクラス</param>
    ' ''' </summary> 
    ' ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ' ''' <remarks>勘定科目マスタから借方細目のコンボボックスをセットする
    ' ''' <para>作成情報：2015/08/11 yu.satoh 
    ' ''' </para></remarks>
    'Public Function CmbKariSaimoku(ByRef Adapter As NpgsqlDataAdapter, ByVal Cn As NpgsqlConnection, ByVal dataEXTM0102 As DataEXTM0102, ByVal karikamokuCd As String)

    '    '開始ログ出力
    '    CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

    '    '変数の宣言
    '    Dim strSQL As String
    '    Dim strWhere As String
    '    Dim strORDER As String

    '    Try

    '        'SQL文(SELECT)
    '        strSQL = strSelectKariSaimoku
    '        'WHERE句
    '        strWhere = " WHERE karikamoku_cd = :karikamokuCd"
    '        'ORDER BY句作成
    '        'strORDER = " ORDER BY karikamoku_cd"
    '        strORDER = " ORDER BY kari_saimoku_cd"

    '        'データアダプタに、SQLを設定
    '        strSQL &= strWhere
    '        strSQL &= strORDER
    '        Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

    '        'バインド変数に型と値をセット
    '        Adapter.SelectCommand.Parameters.Add(New NpgsqlParameter("karikamokuCd", NpgsqlTypes.NpgsqlDbType.Varchar)) '借方科目コード
    '        Adapter.SelectCommand.Parameters("karikamokuCd").Value = karikamokuCd                                       '借方科目コード

    '        '終了ログ出力
    '        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

    '        '正常終了
    '        Return True

    '    Catch ex As Exception
    '        '例外発生
    '        CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
    '        puErrMsg = EXTM0102_E0000 & ex.Message
    '        Return False
    '    End Try

    'End Function

    ' ''' <summary>
    ' ''' 借方内訳コンボボックス作成SQL作成
    ' ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ' ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ' ''' <param name="dataHBKA0101">[IN]データクラス</param>
    ' ''' </summary> 
    ' ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ' ''' <remarks>勘定科目マスタから借方内訳のコンボボックスをセットする
    ' ''' <para>作成情報：2015/08/14 yu.satoh 
    ' ''' </para></remarks>
    'Public Function CmbKariUchi(ByRef Adapter As NpgsqlDataAdapter, ByVal Cn As NpgsqlConnection, ByVal dataEXTM0102 As DataEXTM0102, ByVal karikamokuCd As String, ByVal kariSaimokuCd As String)

    '    '開始ログ出力
    '    CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

    '    '変数の宣言
    '    Dim strSQL As String
    '    Dim strWhere As String
    '    Dim strORDER As String

    '    Try

    '        'SQL文(SELECT)
    '        strSQL = strSelectKariUchi
    '        'WHERE句
    '        strWhere = " WHERE karikamoku_cd = :karikamokuCd" +
    '                   " AND kari_saimoku_cd = :kariSaimokuCd"
    '        'ORDER BY句作成
    '        strORDER = " ORDER BY kari_uchi_cd"

    '        'データアダプタに、SQLを設定
    '        strSQL &= strWhere
    '        strSQL &= strORDER
    '        Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

    '        'バインド変数に型と値をセット
    '        Adapter.SelectCommand.Parameters.Add(New NpgsqlParameter("karikamokuCd", NpgsqlTypes.NpgsqlDbType.Varchar)) '借方科目コード
    '        Adapter.SelectCommand.Parameters("karikamokuCd").Value = karikamokuCd                                       '借方科コード
    '        Adapter.SelectCommand.Parameters.Add(New NpgsqlParameter("kariSaimokuCd", NpgsqlTypes.NpgsqlDbType.Varchar)) '借方細目コード
    '        Adapter.SelectCommand.Parameters("kariSaimokuCd").Value = kariSaimokuCd                                      '借方細目目コード

    '        '終了ログ出力
    '        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

    '        '正常終了
    '        Return True

    '    Catch ex As Exception
    '        '例外発生
    '        CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
    '        puErrMsg = EXTM0102_E0000 & ex.Message
    '        Return False
    '    End Try

    'End Function

    ' ''' <summary>
    ' ''' 借方詳細コンボボックス作成SQL作成
    ' ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ' ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ' ''' <param name="dataHBKA0101">[IN]データクラス</param>
    ' ''' </summary> 
    ' ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ' ''' <remarks>勘定科目マスタから借方詳細のコンボボックスをセットする
    ' ''' <para>作成情報：2015/08/14 yu.satoh 
    ' ''' </para></remarks>
    'Public Function CmbKariShosai(ByRef Adapter As NpgsqlDataAdapter, ByVal Cn As NpgsqlConnection, ByVal dataEXTM0102 As DataEXTM0102, ByVal karikamokuCd As String, ByVal kariSaimokuCd As String, ByVal kariUchiCd As String)

    '    '開始ログ出力
    '    CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

    '    '変数の宣言
    '    Dim strSQL As String
    '    Dim strWhere As String

    '    Try

    '        'SQL文(SELECT)
    '        strSQL = strSelectKariShosai
    '        'WHERE句
    '        strWhere = " WHERE karikamoku_cd = :karikamokuCd" +
    '                   " AND kari_saimoku_cd = :kariSaimokuCd" +
    '                   " AND kari_uchi_cd = :kariUchiCd"

    '        'データアダプタに、SQLを設定
    '        strSQL &= strWhere
    '        Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

    '        'バインド変数に型と値をセット
    '        Adapter.SelectCommand.Parameters.Add(New NpgsqlParameter("karikamokuCd", NpgsqlTypes.NpgsqlDbType.Varchar)) '借方科目コード
    '        Adapter.SelectCommand.Parameters("karikamokuCd").Value = karikamokuCd                                       '借方科目コード
    '        Adapter.SelectCommand.Parameters.Add(New NpgsqlParameter("kariSaimokuCd", NpgsqlTypes.NpgsqlDbType.Varchar)) '借方細目コード
    '        Adapter.SelectCommand.Parameters("kariSaimokuCd").Value = kariSaimokuCd                                      '借方細目コード
    '        Adapter.SelectCommand.Parameters.Add(New NpgsqlParameter("kariUchiCd", NpgsqlTypes.NpgsqlDbType.Varchar))    '借方内訳コード
    '        Adapter.SelectCommand.Parameters("kariUchiCd").Value = kariUchiCd                                            '借方内訳コード

    '        '終了ログ出力
    '        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

    '        '正常終了
    '        Return True

    '    Catch ex As Exception
    '        '例外発生
    '        CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
    '        puErrMsg = EXTM0102_E0000 & ex.Message
    '        Return False
    '    End Try

    'End Function
    ' 2016.04.28 DEL END↑ h.hagiwara レスポンス改善

    ''' <summary>
    ''' 分類表インサートSQL作成
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataHBKA0101">[IN]データクラス</param>
    ''' </summary> 
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>画面の値を、付帯設備分類マスタに登録する。
    ''' <para>作成情報：2015/08/14 yu.satoh 
    ''' </para></remarks>
    Public Function InsertBunrui(ByRef i As Integer, ByRef Adapter As NpgsqlCommand, ByRef Tsx As NpgsqlTransaction, _
                                  ByVal Cn As NpgsqlConnection, ByVal dataEXTM0102 As DataEXTM0102)

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数の宣言
        Dim strSQL As String
        '日付の設定
        Dim dtFrom As String = dataEXTM0102.PropYearFrom.Text + "/" + dataEXTM0102.PropMonthFrom.Text + "/01"
        Dim dtToSub As DateTime = dataEXTM0102.PropYearTo.Text + "/" + dataEXTM0102.PropMonthTo.Text
        '期間ＴＯは、入力された値の月＋１、日付-1の値で最終日を設定できる
        Dim dtTo As String = dtToSub.AddMonths(1).AddDays(-1).ToString("yyyy/MM/dd")

        Try

            'SQL文(INSERT)
            strSQL = strInsertBunrui

            'データアダプタに、SQLを設定
            Adapter = New NpgsqlCommand(strSQL, Cn)

            'バインド変数に型と値をセット
            Adapter.Parameters.Add(New NpgsqlParameter("kikanFrom", NpgsqlTypes.NpgsqlDbType.Varchar))      '期間from
            Adapter.Parameters.Add(New NpgsqlParameter("kikanTo", NpgsqlTypes.NpgsqlDbType.Varchar))        '期間to
            Adapter.Parameters.Add(New NpgsqlParameter("shisetuKbn", NpgsqlTypes.NpgsqlDbType.Varchar))     '施設区分
            Adapter.Parameters.Add(New NpgsqlParameter("bunruiCd", NpgsqlTypes.NpgsqlDbType.Varchar))       '分類コード
            Adapter.Parameters.Add(New NpgsqlParameter("bunruiNm", NpgsqlTypes.NpgsqlDbType.Varchar))       '分類名
            Adapter.Parameters.Add(New NpgsqlParameter("shukeiGrp", NpgsqlTypes.NpgsqlDbType.Varchar))      '集計グループ
            Adapter.Parameters.Add(New NpgsqlParameter("notaxFlg", NpgsqlTypes.NpgsqlDbType.Varchar))       '税
            Adapter.Parameters.Add(New NpgsqlParameter("sort", NpgsqlTypes.NpgsqlDbType.Integer))           '並び順
            Adapter.Parameters.Add(New NpgsqlParameter("kamokuCd", NpgsqlTypes.NpgsqlDbType.Varchar))       '科目コード
            Adapter.Parameters.Add(New NpgsqlParameter("saimokuCd", NpgsqlTypes.NpgsqlDbType.Varchar))      '細目コード
            Adapter.Parameters.Add(New NpgsqlParameter("uchiCd", NpgsqlTypes.NpgsqlDbType.Varchar))         '内訳コード
            Adapter.Parameters.Add(New NpgsqlParameter("shosaiCd", NpgsqlTypes.NpgsqlDbType.Varchar))       '詳細コード
            Adapter.Parameters.Add(New NpgsqlParameter("karikamokuCd", NpgsqlTypes.NpgsqlDbType.Varchar))   '借方科目コード
            Adapter.Parameters.Add(New NpgsqlParameter("karisaimokuCd", NpgsqlTypes.NpgsqlDbType.Varchar))  '借方細目コード
            Adapter.Parameters.Add(New NpgsqlParameter("kariuchiCd", NpgsqlTypes.NpgsqlDbType.Varchar))     '借方内訳コード
            Adapter.Parameters.Add(New NpgsqlParameter("karishosaiCd", NpgsqlTypes.NpgsqlDbType.Varchar))   '借方詳細コード
            Adapter.Parameters.Add(New NpgsqlParameter("sts", NpgsqlTypes.NpgsqlDbType.Varchar))            '無効フラグ
            Adapter.Parameters.Add(New NpgsqlParameter("RIYORYO_FLG", NpgsqlTypes.NpgsqlDbType.Varchar))    ' 付帯利用料フラグ       ' 2015.11.30 ADD h.hagiwara
            Adapter.Parameters.Add(New NpgsqlParameter("addDate", NpgsqlTypes.NpgsqlDbType.Timestamp))      '作成日付
            Adapter.Parameters.Add(New NpgsqlParameter("addUserCd", NpgsqlTypes.NpgsqlDbType.Varchar))      '作成者名
            Adapter.Parameters.Add(New NpgsqlParameter("upDate", NpgsqlTypes.NpgsqlDbType.Timestamp))       '更新日
            Adapter.Parameters.Add(New NpgsqlParameter("upUserCd", NpgsqlTypes.NpgsqlDbType.Varchar))       '更新者

            '値をセット
            Adapter.Parameters("kikanFrom").Value = dtFrom  '日付from
            Adapter.Parameters("kikanTo").Value = dtTo      '日付to
            '施設区分がシアターなら1、スタジオなら2を設定
            If dataEXTM0102.PropTheaterBtn.Checked = True Then
                Adapter.Parameters("shisetuKbn").Value = "1"
            Else
                Adapter.Parameters("shisetuKbn").Value = "2"
            End If
            'Adapter.Parameters("bunruiCd").Value = dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 0).Value '分類コード
            Adapter.Parameters("bunruiCd").Value = Format(Integer.Parse(dataEXTM0102.PropMaxBunruiCd), "00")
            Adapter.Parameters("bunruiNm").Value = dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 1).Value '分類名
            Adapter.Parameters("shukeiGrp").Value = dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 2).Value '集計グループ
            '税フラグがチェックされていれば1、されていなければ0を設定
            If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 3).Value = True Then
                Adapter.Parameters("notaxFlg").Value = "1"
            Else
                Adapter.Parameters("notaxFlg").Value = "0"
            End If
            ' 2016.04.28 UPD START↓ h.hagiwara レスポンス改善
            'Adapter.Parameters("sort").Value = dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 12).Value   '並び順
            Adapter.Parameters("sort").Value = dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 8).Value   '並び順
            ' 2016.04.28 UPD END↑ h.hagiwara レスポンス改善
            Adapter.Parameters("kamokuCd").Value = dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 4).Value '科目コード
            Adapter.Parameters("saimokuCd").Value = dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 5).Value '細目コード
            Adapter.Parameters("uchiCd").Value = dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 6).Value    '内訳コード
            Adapter.Parameters("shosaiCd").Value = dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 7).Value  '細目コード
            ' 2016.04.28 UPD START↓ h.hagiwara レスポンス改善
            'Adapter.Parameters("karikamokuCd").Value = dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 8).Value '借方科目コード
            'Adapter.Parameters("karisaimokuCd").Value = dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 9).Value '借方細目コード
            'Adapter.Parameters("kariuchiCd").Value = dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 10).Value   '借方内訳コード
            'Adapter.Parameters("karishosaiCd").Value = dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 11).Value '借方詳細コード
            ''無効フラグがチェックされていれば1を、そうでなければ0を設定
            'If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 13).Value = True Then
            '    Adapter.Parameters("sts").Value = "1"
            'Else
            '    Adapter.Parameters("sts").Value = "0"
            'End If
            '' 2015.11.30 ADD STATR↓ h.hagiwara
            '' 付帯利用料フラグがチェックされていれば1を、そうでなければ0を設定
            'If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 14).Value = True Then
            '    Adapter.Parameters("RIYORYO_FLG").Value = "1"
            'Else
            '    Adapter.Parameters("RIYORYO_FLG").Value = "0"
            'End If
            '' 2015.11.30 ADD END↑ h.hagiwara
            Adapter.Parameters("karikamokuCd").Value = DBNull.Value    '借方科目コード
            Adapter.Parameters("karisaimokuCd").Value = DBNull.Value   '借方細目コード
            Adapter.Parameters("kariuchiCd").Value = DBNull.Value      '借方内訳コード
            Adapter.Parameters("karishosaiCd").Value = DBNull.Value    '借方詳細コード
            '無効フラグがチェックされていれば1を、そうでなければ0を設定
            If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 9).Value = True Then
                Adapter.Parameters("sts").Value = "1"
            Else
                Adapter.Parameters("sts").Value = "0"
            End If
            ' 付帯利用料フラグがチェックされていれば1を、そうでなければ0を設定
            If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 10).Value = True Then
                Adapter.Parameters("RIYORYO_FLG").Value = "1"
            Else
                Adapter.Parameters("RIYORYO_FLG").Value = "0"
            End If
            ' 2016.04.28 UPD END↑ h.hagiwara レスポンス改善

            '作成、更新日、及び作成者、作成日は、現在日付を設定。UserCdは基本的にログインユーザーのコード
            Adapter.Parameters("addDate").Value = Now
            'Adapter.Parameters("addUserCd").Value = "testIns"
            Adapter.Parameters("addUserCd").Value = CommonEXT.PropComStrUserId
            Adapter.Parameters("upDate").Value = Now
            'Adapter.Parameters("upUserCd").Value = "testIns"
            Adapter.Parameters("upUserCd").Value = CommonEXT.PropComStrUserId

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter)
            puErrMsg = EXTM0102_E0000 & ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' 分類表アップデートSQL作成
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataHBKA0101">[IN]データクラス</param>
    ''' </summary> 
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>画面の値を、付帯設備分類マスタに登録する。（更新）
    ''' <para>作成情報：2015/08/14 yu.satoh 
    ''' </para></remarks>
    Public Function UpdateBunrui(ByRef i As Integer, ByRef Adapter As NpgsqlCommand, ByRef Tsx As NpgsqlTransaction, _
                                  ByVal Cn As NpgsqlConnection, ByVal dataEXTM0102 As DataEXTM0102)

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)


        '変数の宣言
        Dim strSQL As String
        Dim strWhere As String
        Dim dtDate As String = Now

        Try

            'SQL文(UPDATE)
            strSQL = strUpdateBunrui

            'where句作成
            strWhere = " WHERE kikan_from = :kikanFromUp"
            strWhere &= " AND kikan_to = :kikanToUp"
            strWhere &= " AND bunrui_cd = :bunruiCd"
            strWhere &= " AND shisetu_kbn = :shisetuKbn"

            'データアダプタに、SQLを設定
            strSQL &= strWhere
            Adapter = New NpgsqlCommand(strSQL, Cn)

            '日付を設定。fromは画面from月の初めを、toは画面toの月を+1し、その日付-1をして最終日を設定
            Dim dtFrom As String = dataEXTM0102.PropYearFrom.Text + "/" + dataEXTM0102.PropMonthFrom.Text + "/01"
            Dim dtToSub As DateTime = dataEXTM0102.PropYearTo.Text + "/" + dataEXTM0102.PropMonthTo.Text
            Dim dtTo As String = dtToSub.AddMonths(1).AddDays(-1).ToString("yyyy/MM/dd")
            '日付を設定。fromは画面from月の初めを、toは画面toの月を+1し、その日付-1をして最終日を設定
            Dim dtFromUp As String = dataEXTM0102.PropFromYearUp + "/" + dataEXTM0102.PropFromMonthUp + "/01"
            Dim dtToSubUp As DateTime = dataEXTM0102.PropToYearUp + "/" + dataEXTM0102.PropToMonthUp
            Dim dtToUp As String = dtToSubUp.AddMonths(1).AddDays(-1).ToString("yyyy/MM/dd")

            'dataEXTM0102.PropMaxBunruiCd = dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 0).Value

            'バインド変数に型と値をセット
            Adapter.Parameters.Add(New NpgsqlParameter("kikanFrom", NpgsqlTypes.NpgsqlDbType.Varchar)) '期間from
            Adapter.Parameters.Add(New NpgsqlParameter("kikanTo", NpgsqlTypes.NpgsqlDbType.Varchar))   '期間to
            Adapter.Parameters.Add(New NpgsqlParameter("shisetuKbn", NpgsqlTypes.NpgsqlDbType.Varchar)) '施設区分
            Adapter.Parameters.Add(New NpgsqlParameter("bunruiCd", NpgsqlTypes.NpgsqlDbType.Varchar))   '分類コード
            Adapter.Parameters.Add(New NpgsqlParameter("bunruiNm", NpgsqlTypes.NpgsqlDbType.Varchar))   '分類名
            Adapter.Parameters.Add(New NpgsqlParameter("shukeiGrp", NpgsqlTypes.NpgsqlDbType.Varchar))  '集計グループ
            Adapter.Parameters.Add(New NpgsqlParameter("notaxFlg", NpgsqlTypes.NpgsqlDbType.Varchar))   '税
            Adapter.Parameters.Add(New NpgsqlParameter("sort", NpgsqlTypes.NpgsqlDbType.Integer))       '並び順
            Adapter.Parameters.Add(New NpgsqlParameter("kamokuCd", NpgsqlTypes.NpgsqlDbType.Varchar))   '科目コード
            Adapter.Parameters.Add(New NpgsqlParameter("saimokuCd", NpgsqlTypes.NpgsqlDbType.Varchar))  '細目コード
            Adapter.Parameters.Add(New NpgsqlParameter("uchiCd", NpgsqlTypes.NpgsqlDbType.Varchar))     '内訳コード
            Adapter.Parameters.Add(New NpgsqlParameter("shosaiCd", NpgsqlTypes.NpgsqlDbType.Varchar))   '詳細コード
            Adapter.Parameters.Add(New NpgsqlParameter("karikamokuCd", NpgsqlTypes.NpgsqlDbType.Varchar)) '借方科目コード
            Adapter.Parameters.Add(New NpgsqlParameter("karisaimokuCd", NpgsqlTypes.NpgsqlDbType.Varchar)) '借方細目コード
            Adapter.Parameters.Add(New NpgsqlParameter("kariuchiCd", NpgsqlTypes.NpgsqlDbType.Varchar))     '借方内訳コード
            Adapter.Parameters.Add(New NpgsqlParameter("karishosaiCd", NpgsqlTypes.NpgsqlDbType.Varchar))   '借方詳細コード
            Adapter.Parameters.Add(New NpgsqlParameter("sts", NpgsqlTypes.NpgsqlDbType.Varchar))            '無効フラグ
            Adapter.Parameters.Add(New NpgsqlParameter("RIYORYO_FLG", NpgsqlTypes.NpgsqlDbType.Varchar))    ' 付帯利用料フラグ       ' 2015.11.30 ADD h.hagiwara
            Adapter.Parameters.Add(New NpgsqlParameter("upDate", NpgsqlTypes.NpgsqlDbType.Timestamp))       '更新日
            Adapter.Parameters.Add(New NpgsqlParameter("upUserCd", NpgsqlTypes.NpgsqlDbType.Varchar))       '更新者
            Adapter.Parameters.Add(New NpgsqlParameter("kikanFromUp", NpgsqlTypes.NpgsqlDbType.Varchar))    '変更前期間from
            Adapter.Parameters.Add(New NpgsqlParameter("kikanToUp", NpgsqlTypes.NpgsqlDbType.Varchar))      '変更前期間to

            '値をセット
            '日付
            Adapter.Parameters("kikanFrom").Value = dtFrom
            Adapter.Parameters("kikanTo").Value = dtTo
            '施設区分がシアターなら1、スタジオなら2を設定
            If dataEXTM0102.PropTheaterBtn.Checked = True Then
                Adapter.Parameters("shisetuKbn").Value = "1"
            Else
                Adapter.Parameters("shisetuKbn").Value = "2"
            End If
            Adapter.Parameters("bunruiCd").Value = dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 0).Value '分類コード
            Adapter.Parameters("bunruiNm").Value = dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 1).Value '分類名
            Adapter.Parameters("shukeiGrp").Value = dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 2).Value '集計グループ
            '税フラグがチェックされていれば1、されていなければ0を設定
            If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 3).Value = True Then
                Adapter.Parameters("notaxFlg").Value = "1"
            Else
                Adapter.Parameters("notaxFlg").Value = "0"
            End If
            ' 2016.04.28 UPD START↓ h.hagiwara レスポンス改善
            'Adapter.Parameters("sort").Value = dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 12).Value      '並び順
            Adapter.Parameters("sort").Value = dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 8).Value      '並び順
            ' 2016.04.28 UPD END↑ h.hagiwara レスポンス改善
            Adapter.Parameters("kamokuCd").Value = dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 4).Value   '科目コード
            Adapter.Parameters("saimokuCd").Value = dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 5).Value  '細目コード
            Adapter.Parameters("uchiCd").Value = dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 6).Value     '内訳コード
            Adapter.Parameters("shosaiCd").Value = dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 7).Value   '詳細コード
            ' 2016.04.28 UPD START↓ h.hagiwara レスポンス改善
            'Adapter.Parameters("karikamokuCd").Value = dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 8).Value '借方科目コード
            'Adapter.Parameters("karisaimokuCd").Value = dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 9).Value '借方細目コード
            'Adapter.Parameters("kariuchiCd").Value = dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 10).Value    '借方内訳コード
            'Adapter.Parameters("karishosaiCd").Value = dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 11).Value  '借方詳細コード
            ''無効フラグがチェックされていれば1を、そうでなければ0を設定
            'If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 13).Value = True Then
            '    Adapter.Parameters("sts").Value = "1"
            'Else
            '    Adapter.Parameters("sts").Value = "0"
            'End If
            '' 2015.11.30 ADD STATR↓ h.hagiwara
            '' 付帯利用料フラグがチェックされていれば1を、そうでなければ0を設定
            'If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 14).Value = True Then
            '    Adapter.Parameters("RIYORYO_FLG").Value = "1"
            'Else
            '    Adapter.Parameters("RIYORYO_FLG").Value = "0"
            'End If
            '' 2015.11.30 ADD END↑ h.hagiwara
            Adapter.Parameters("karikamokuCd").Value = DBNull.Value        '借方科目コード
            Adapter.Parameters("karisaimokuCd").Value = DBNull.Value       '借方細目コード
            Adapter.Parameters("kariuchiCd").Value = DBNull.Value          '借方内訳コード
            Adapter.Parameters("karishosaiCd").Value = DBNull.Value        '借方詳細コード
            '無効フラグがチェックされていれば1を、そうでなければ0を設定
            If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 9).Value = True Then
                Adapter.Parameters("sts").Value = "1"
            Else
                Adapter.Parameters("sts").Value = "0"
            End If
            ' 付帯利用料フラグがチェックされていれば1を、そうでなければ0を設定
            If dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 10).Value = True Then
                Adapter.Parameters("RIYORYO_FLG").Value = "1"
            Else
                Adapter.Parameters("RIYORYO_FLG").Value = "0"
            End If
            ' 2016.04.28 UPD END↑ h.hagiwara レスポンス改善

            '更新日、及び更新者は、現在日付を設定。UserCdは基本的にログインユーザーのコード
            Adapter.Parameters("upDate").Value = Now
            'Adapter.Parameters("upUserCd").Value = "test"
            Adapter.Parameters("upUserCd").Value = CommonEXT.PropComStrUserId
            '日付
            Adapter.Parameters("kikanFromUp").Value = dtFromUp
            Adapter.Parameters("kikanToUp").Value = dtToUp

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter)
            puErrMsg = EXTM0102_E0000 & ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' 付帯表インサートSQL作成
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataHBKA0101">[IN]データクラス</param>
    ''' </summary> 
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>画面の値を、付帯設備マスタに登録する。
    ''' <para>作成情報：2015/08/14 yu.satoh 
    ''' </para></remarks>
    Public Function InsertFutai(ByRef i As Integer, ByRef Adapter As NpgsqlCommand, ByRef Tsx As NpgsqlTransaction, _
                                ByVal Cn As NpgsqlConnection, ByVal dataEXTM0102 As DataEXTM0102)

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数の宣言
        Dim strSQL As String

        Try

            'SQL文(INSERT)
            strSQL = strInsertFutai

            'データアダプタに、SQLを設定
            Adapter = New NpgsqlCommand(strSQL, Cn)

            '日付を設定。fromは画面from月の初めを、toは画面toの月を+1し、その日付-1をして最終日を設定
            Dim dtFrom As String = dataEXTM0102.PropYearFrom.Text + "/" + dataEXTM0102.PropMonthFrom.Text + "/01"
            Dim dtToSub As DateTime = dataEXTM0102.PropYearTo.Text + "/" + dataEXTM0102.PropMonthTo.Text
            Dim dtTo As String = dtToSub.AddMonths(1).AddDays(-1).ToString("yyyy/MM/dd")
            dataEXTM0102.PropMaxFutaiBunruiCd = dataEXTM0102.PropMaxFutaiBunruiCd + 1

            'バインド変数に型と値をセット
            Adapter.Parameters.Add(New NpgsqlParameter("kikanFrom", NpgsqlTypes.NpgsqlDbType.Varchar))  '期間from
            Adapter.Parameters.Add(New NpgsqlParameter("kikanTo", NpgsqlTypes.NpgsqlDbType.Varchar))    '期間to
            Adapter.Parameters.Add(New NpgsqlParameter("shisetuKbn", NpgsqlTypes.NpgsqlDbType.Varchar)) '施設区分
            Adapter.Parameters.Add(New NpgsqlParameter("bunruiCd", NpgsqlTypes.NpgsqlDbType.Varchar))   '分類コード
            Adapter.Parameters.Add(New NpgsqlParameter("futaiCd", NpgsqlTypes.NpgsqlDbType.Varchar))    '付帯コード
            Adapter.Parameters.Add(New NpgsqlParameter("futaiNm", NpgsqlTypes.NpgsqlDbType.Varchar))    '付帯名
            Adapter.Parameters.Add(New NpgsqlParameter("tanka", NpgsqlTypes.NpgsqlDbType.Integer))      '単価
            Adapter.Parameters.Add(New NpgsqlParameter("tani", NpgsqlTypes.NpgsqlDbType.Varchar))       '単位
            Adapter.Parameters.Add(New NpgsqlParameter("defFlg", NpgsqlTypes.NpgsqlDbType.Varchar))     'デフォルト
            Adapter.Parameters.Add(New NpgsqlParameter("sort", NpgsqlTypes.NpgsqlDbType.Integer))       '並び順
            Adapter.Parameters.Add(New NpgsqlParameter("sts", NpgsqlTypes.NpgsqlDbType.Varchar))        '無効フラグ
            Adapter.Parameters.Add(New NpgsqlParameter("addDate", NpgsqlTypes.NpgsqlDbType.Timestamp))  '作成日
            Adapter.Parameters.Add(New NpgsqlParameter("addUserCd", NpgsqlTypes.NpgsqlDbType.Varchar))  '作成者
            Adapter.Parameters.Add(New NpgsqlParameter("upDate", NpgsqlTypes.NpgsqlDbType.Timestamp))   '更新日
            Adapter.Parameters.Add(New NpgsqlParameter("upUserCd", NpgsqlTypes.NpgsqlDbType.Varchar))   '更新者

            '値をセット
            '日付
            Adapter.Parameters("kikanFrom").Value = dtFrom
            Adapter.Parameters("kikanTo").Value = dtTo
            '施設区分がシアターなら1、スタジオなら2を設定
            If dataEXTM0102.PropTheaterBtn.Checked = True Then
                Adapter.Parameters("shisetuKbn").Value = "1"
            Else
                Adapter.Parameters("shisetuKbn").Value = "2"
            End If
            'Adapter.Parameters("bunruiCd").Value = dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 0).Value '分類コード
            'Adapter.Parameters("futaiCd").Value = dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, 0).Value    '付帯コード
            'Adapter.Parameters("bunruiCd").Value = Format(Integer.Parse(dataEXTM0102.PropUpdBunruiCd), "00")
            'Adapter.Parameters("bunruiCd").Value = Format(Integer.Parse(dataEXTM0102.PropBunruiCd), "00")
            'Adapter.Parameters("futaiCd").Value = Format(dataEXTM0102.PropMaxFutaiBunruiCd, "000")
            'Adapter.Parameters("futaiNm").Value = dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, 1).Value    '付帯名
            'Adapter.Parameters("tanka").Value = dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, 2).Value      '単価
            'Adapter.Parameters("tani").Value = dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, 3).Value       '単位
            ''デフォルトがチェックされていれば1、されていなければ0
            'If dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, 4).Value = True Then
            '    Adapter.Parameters("defFlg").Value = "1"
            'Else
            '    Adapter.Parameters("defFlg").Value = "0"
            'End If

            'Adapter.Parameters("sort").Value = dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, 5).Value       '並び順

            ''無効フラグがチェックされていれば1を、そうでなければ0を設定
            'If dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, 6).Value = True Then
            '    Adapter.Parameters("sts").Value = "1"
            'Else
            '    Adapter.Parameters("sts").Value = "0"
            'End If
            Adapter.Parameters("bunruiCd").Value = Format(Integer.Parse(dataEXTM0102.PropBunruiCd), "00")
            Adapter.Parameters("futaiCd").Value = Format(dataEXTM0102.PropMaxFutaiBunruiCd, "000")
            Adapter.Parameters("futaiNm").Value = dataEXTM0102.PropDtFutaiMst.Rows(i).Item(1)   '付帯名
            Adapter.Parameters("tanka").Value = dataEXTM0102.PropDtFutaiMst.Rows(i).Item(2)      '単価
            Adapter.Parameters("tani").Value = dataEXTM0102.PropDtFutaiMst.Rows(i).Item(3)       '単位
            Adapter.Parameters("defFlg").Value = dataEXTM0102.PropDtFutaiMst.Rows(i).Item(4)
            Adapter.Parameters("sort").Value = dataEXTM0102.PropDtFutaiMst.Rows(i).Item(5)      '並び順
            Adapter.Parameters("sts").Value = dataEXTM0102.PropDtFutaiMst.Rows(i).Item(6)
            '作成者、作成日、更新者、更新日は現在日付を設定。UserCdは基本的にログインユーザーのコード
            Adapter.Parameters("addDate").Value = Now
            'Adapter.Parameters("addUserCd").Value = "testIns"
            Adapter.Parameters("addUserCd").Value = CommonEXT.PropComStrUserId
            Adapter.Parameters("upDate").Value = Now
            'Adapter.Parameters("upUserCd").Value = "testIns"
            Adapter.Parameters("upUserCd").Value = CommonEXT.PropComStrUserId

            'コマンド実行
            Adapter.ExecuteNonQuery()

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter)
            puErrMsg = EXTM0102_E0000 & ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' 付帯設備表アップデートSQL作成
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataHBKA0101">[IN]データクラス</param>
    ''' </summary> 
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>画面の値を、付帯設備マスタに登録する。（更新）
    ''' <para>作成情報：2015/08/14 yu.satoh 
    ''' </para></remarks>
    Public Function UpdateFutai(ByRef i As Integer, ByRef Adapter As NpgsqlCommand, ByRef Tsx As NpgsqlTransaction, _
                                ByVal Cn As NpgsqlConnection, ByVal dataEXTM0102 As DataEXTM0102)

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)


        '変数の宣言
        Dim strSQL As String
        Dim strWhere As String
        '付帯コードの文字数チェック
        'Dim lenFutai As Integer = dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, 0).Value.ToString.Length

        Try

            'SQL文(UPDATE)
            strSQL = strUpdateFutai

            'where句作成
            strWhere = " WHERE kikan_from = :kikanFromUp"
            strWhere &= " AND kikan_to = :kikanToUp"
            strWhere &= " AND futai_cd = :futaiCd"
            strWhere &= " AND bunrui_cd = :bunruiCd"
            strWhere &= " AND shisetu_kbn = :shisetuKbn"

            'データアダプタに、SQLを設定
            strSQL &= strWhere
            Adapter = New NpgsqlCommand(strSQL, Cn)

            '日付を設定。fromは画面from月の初めを、toは画面toの月を+1し、その日付-1をして最終日を設定
            Dim dtFrom As String = dataEXTM0102.PropYearFrom.Text + "/" + dataEXTM0102.PropMonthFrom.Text + "/01"
            Dim dtToSub As DateTime = dataEXTM0102.PropYearTo.Text + "/" + dataEXTM0102.PropMonthTo.Text
            Dim dtTo As String = dtToSub.AddMonths(1).AddDays(-1).ToString("yyyy/MM/dd")
            '日付を設定。fromは画面from月の初めを、toは画面toの月を+1し、その日付-1をして最終日を設定
            Dim dtFromUp As String = dataEXTM0102.PropFromYearUp + "/" + dataEXTM0102.PropFromMonthUp + "/01"
            Dim dtToSubUp As DateTime = dataEXTM0102.PropToYearUp + "/" + dataEXTM0102.PropToMonthUp
            Dim dtToUp As String = dtToSubUp.AddMonths(1).AddDays(-1).ToString("yyyy/MM/dd")

            'バインド変数に型と値をセット
            Adapter.Parameters.Add(New NpgsqlParameter("kikanFrom", NpgsqlTypes.NpgsqlDbType.Varchar))  '期間from
            Adapter.Parameters.Add(New NpgsqlParameter("kikanTo", NpgsqlTypes.NpgsqlDbType.Varchar))    '期間to
            Adapter.Parameters.Add(New NpgsqlParameter("shisetuKbn", NpgsqlTypes.NpgsqlDbType.Varchar)) '施設区分
            Adapter.Parameters.Add(New NpgsqlParameter("bunruiCd", NpgsqlTypes.NpgsqlDbType.Varchar))   '分類コード
            Adapter.Parameters.Add(New NpgsqlParameter("futaiCd", NpgsqlTypes.NpgsqlDbType.Varchar))    '付帯コード
            Adapter.Parameters.Add(New NpgsqlParameter("futaiNm", NpgsqlTypes.NpgsqlDbType.Varchar))    '付帯名
            Adapter.Parameters.Add(New NpgsqlParameter("tanka", NpgsqlTypes.NpgsqlDbType.Integer))      '単価
            Adapter.Parameters.Add(New NpgsqlParameter("tani", NpgsqlTypes.NpgsqlDbType.Varchar))       '単位
            Adapter.Parameters.Add(New NpgsqlParameter("defFlg", NpgsqlTypes.NpgsqlDbType.Varchar))     'デフォルト
            Adapter.Parameters.Add(New NpgsqlParameter("sort", NpgsqlTypes.NpgsqlDbType.Integer))       '並び順
            Adapter.Parameters.Add(New NpgsqlParameter("sts", NpgsqlTypes.NpgsqlDbType.Varchar))        '無効フラグ
            Adapter.Parameters.Add(New NpgsqlParameter("upDate", NpgsqlTypes.NpgsqlDbType.Timestamp))   '更新日
            Adapter.Parameters.Add(New NpgsqlParameter("upUserCd", NpgsqlTypes.NpgsqlDbType.Varchar))   '更新者
            Adapter.Parameters.Add(New NpgsqlParameter("kikanFromUp", NpgsqlTypes.NpgsqlDbType.Varchar))    '変更前期間from
            Adapter.Parameters.Add(New NpgsqlParameter("kikanToUp", NpgsqlTypes.NpgsqlDbType.Varchar))      '変更前期間to

            '値をセット
            '日付
            Adapter.Parameters("kikanFrom").Value = dtFrom
            Adapter.Parameters("kikanTo").Value = dtTo
            '施設区分がシアターなら1、スタジオなら2を設定
            If dataEXTM0102.PropTheaterBtn.Checked = True Then
                Adapter.Parameters("shisetuKbn").Value = "1"
            Else
                Adapter.Parameters("shisetuKbn").Value = "2"
            End If
            'Adapter.Parameters("bunruiCd").Value = dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 0).Value '分類コード
            'Adapter.Parameters("bunruiCd").Value = Format(Integer.Parse(dataEXTM0102.PropUpdBunruiCd), "00")
            'Adapter.Parameters("futaiCd").Value = dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, 0).Value     '付帯コード
            'Adapter.Parameters("futaiNm").Value = dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, 1).Value     '付帯名
            'Adapter.Parameters("tanka").Value = dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, 2).Value       '単価
            'Adapter.Parameters("tani").Value = dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, 3).Value        '単位
            ''デフォルトがチェックされていれば1、されていなければ0
            'If dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, 4).Value = True Then
            '    Adapter.Parameters("defFlg").Value = "1"
            'Else
            '    Adapter.Parameters("defFlg").Value = "0"
            'End If

            'Adapter.Parameters("sort").Value = dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, 5).Value      '並び順

            ''無効フラグがチェックされていれば1を、そうでなければ0を設定
            'If dataEXTM0102.PropVwFutaiSheet.ActiveSheet.Cells(i, 6).Value = True Then
            '    Adapter.Parameters("sts").Value = "1"
            'Else
            '    Adapter.Parameters("sts").Value = "0"
            'End If
            Adapter.Parameters("bunruiCd").Value = Format(Integer.Parse(dataEXTM0102.PropBunruiCd), "00")
            Adapter.Parameters("futaiCd").Value = Format(Integer.Parse(dataEXTM0102.PropDtFutaiMst.Rows(i).Item(0)), "000")
            Adapter.Parameters("futaiNm").Value = dataEXTM0102.PropDtFutaiMst.Rows(i).Item(1)   '付帯名
            Adapter.Parameters("tanka").Value = dataEXTM0102.PropDtFutaiMst.Rows(i).Item(2)      '単価
            Adapter.Parameters("tani").Value = dataEXTM0102.PropDtFutaiMst.Rows(i).Item(3)       '単位
            Adapter.Parameters("defFlg").Value = dataEXTM0102.PropDtFutaiMst.Rows(i).Item(4)
            Adapter.Parameters("sort").Value = dataEXTM0102.PropDtFutaiMst.Rows(i).Item(5)      '並び順
            Adapter.Parameters("sts").Value = dataEXTM0102.PropDtFutaiMst.Rows(i).Item(6)

            '更新日、更新者は、現在日付を設定。UserCdは基本的にログインユーザーのコード
            Adapter.Parameters("upDate").Value = Now
            'Adapter.Parameters("upUserCd").Value = "testIns"
            Adapter.Parameters("upUserCd").Value = CommonEXT.PropComStrUserId

            '日付
            Adapter.Parameters("kikanFromUp").Value = dtFromUp
            Adapter.Parameters("kikanToUp").Value = dtToUp

            'コマンド実行
            Adapter.ExecuteNonQuery()

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter)
            puErrMsg = EXTM0102_E0000 & ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' 付属設備品分類マスタの最大分類コード取得SQL作成
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataEXTM0102">[IN]データクラス</param>
    ''' </summary> 
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>付属設備品分類マスタから最大分類コードを取得する
    ''' <para>作成情報：2015/08/11 yu.satoh 
    ''' </para></remarks>
    Public Function SelectBunruiCD(ByRef Adapter As NpgsqlDataAdapter, ByVal Cn As NpgsqlConnection, ByVal dataEXTM0102 As DataEXTM0102)

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数の宣言
        Dim strSQL As String
        Dim strWhere As String
        Dim dtFromUp As String
        Dim dtToUp As String

        Try

            'SQL文(SELECT)
            strSQL = strSelectbunricd
            'where句作成
            strWhere = " WHERE kikan_from = :kikanFrom"
            strWhere &= " AND kikan_to = :kikanTo"
            strWhere &= " AND shisetu_kbn = :shisetuKbn"
            strWhere &= " AND bunrui_cd <> '99' "


            'データアダプタに、SQLを設定
            strSQL &= strWhere

            'データアダプタに、SQLを設定
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            '日付を設定。fromは画面from月の初めを、toは画面toの月を+1し、その日付-1をして最終日を設定
            'Dim dtFrom As String = dataEXTM0102.PropYearFrom.Text + "/" + dataEXTM0102.PropMonthFrom.Text + "/01"
            'Dim dtToSub As DateTime = dataEXTM0102.PropYearTo.Text + "/" + dataEXTM0102.PropMonthTo.Text
            Dim dtToSub As DateTime
            'Dim dtTo As String = dtToSub.AddMonths(1).AddDays(-1).ToString("yyyy/MM/dd")
            '日付を設定。fromは画面from月の初めを、toは画面toの月を+1し、その日付-1をして最終日を設定
            If dataEXTM0102.PropNewBtn.Checked = True Then
                'dtFromUp = dtFrom
                'dtToUp = dtTo
                dtFromUp = dataEXTM0102.PropYearFrom.Text + "/" + dataEXTM0102.PropMonthFrom.Text + "/01"
                dtToSub = dataEXTM0102.PropYearTo.Text + "/" + dataEXTM0102.PropMonthTo.Text
                dtToUp = dtToSub.AddMonths(1).AddDays(-1).ToString("yyyy/MM/dd")
            Else
                dtFromUp = dataEXTM0102.PropFromYearUp + "/" + dataEXTM0102.PropFromMonthUp + "/01"
                Dim dtToSubUp As DateTime = dataEXTM0102.PropToYearUp + "/" + dataEXTM0102.PropToMonthUp
                dtToUp = (dtToSubUp.AddMonths(1).AddDays(-1).ToString("yyyy/MM/dd"))
            End If

            'バインド変数に型と値をセット
            Adapter.SelectCommand.Parameters.Add(New NpgsqlParameter("kikanFrom", NpgsqlTypes.NpgsqlDbType.Varchar)) '期間from
            Adapter.SelectCommand.Parameters("kikanFrom").Value = dtFromUp                                           '期間from
            Adapter.SelectCommand.Parameters.Add(New NpgsqlParameter("kikanTo", NpgsqlTypes.NpgsqlDbType.Varchar))   '期間to
            Adapter.SelectCommand.Parameters("kikanTo").Value = dtToUp                                             '期間from
            Adapter.SelectCommand.Parameters.Add(New NpgsqlParameter("shisetuKbn", NpgsqlTypes.NpgsqlDbType.Varchar)) '施設区分
            If dataEXTM0102.PropTheaterBtn.Checked = True Then
                Adapter.SelectCommand.Parameters("shisetuKbn").Value = "1"
            Else
                Adapter.SelectCommand.Parameters("shisetuKbn").Value = "2"
            End If

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
            puErrMsg = EXTM0102_E0000 & ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' 付属設備品分類マスタの最大分類コード取得SQL作成
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataEXTM0102">[IN]データクラス</param>
    ''' </summary> 
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>付属設備品分類マスタから最大分類コードを取得する
    ''' <para>作成情報：2015/08/11 yu.satoh 
    ''' </para></remarks>
    Public Function SelectFutaiBunruiCD(ByRef Adapter As NpgsqlDataAdapter, ByVal Cn As NpgsqlConnection, ByVal dataEXTM0102 As DataEXTM0102)

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数の宣言
        Dim strSQL As String
        Dim strWhere As String
        Dim dtFromUp As String
        Dim dtToUp As String

        Try

            'SQL文(SELECT)
            strSQL = strSelectfutaibunricd
            'where句作成
            strWhere = " WHERE kikan_from = :kikanFrom"
            strWhere &= " AND kikan_to = :kikanTo"
            strWhere &= " AND bunrui_cd = :bunruiCd"
            strWhere &= " AND shisetu_kbn = :shisetuKbn"

            'データアダプタに、SQLを設定
            strSQL &= strWhere
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            '日付を設定。fromは画面from月の初めを、toは画面toの月を+1し、その日付-1をして最終日を設定
            Dim dtFrom As String = dataEXTM0102.PropYearFrom.Text + "/" + dataEXTM0102.PropMonthFrom.Text + "/01"
            Dim dtToSub As DateTime = dataEXTM0102.PropYearTo.Text + "/" + dataEXTM0102.PropMonthTo.Text
            Dim dtTo As String = dtToSub.AddMonths(1).AddDays(-1).ToString("yyyy/MM/dd")
            '日付を設定。fromは画面from月の初めを、toは画面toの月を+1し、その日付-1をして最終日を設定
            If dataEXTM0102.PropFromYearUp = "" Or dataEXTM0102.PropFromYearUp = Nothing Then
                dtFromUp = dtFrom
                dtToUp = dtTo
            Else
                dtFromUp = dataEXTM0102.PropFromYearUp + "/" + dataEXTM0102.PropFromMonthUp + "/01"
                Dim dtToSubUp As DateTime = dataEXTM0102.PropToYearUp + "/" + dataEXTM0102.PropToMonthUp
                dtToUp = (dtToSubUp.AddMonths(1).AddDays(-1).ToString("yyyy/MM/dd"))
            End If

            'バインド変数に型と値をセット
            Adapter.SelectCommand.Parameters.Add(New NpgsqlParameter("kikanFrom", NpgsqlTypes.NpgsqlDbType.Varchar)) '期間from
            Adapter.SelectCommand.Parameters("kikanFrom").Value = dtFromUp                                           '期間from
            Adapter.SelectCommand.Parameters.Add(New NpgsqlParameter("kikanTo", NpgsqlTypes.NpgsqlDbType.Varchar))   '期間to
            Adapter.SelectCommand.Parameters("kikanTo").Value = dtToUp                                             '期間from
            Adapter.SelectCommand.Parameters.Add(New NpgsqlParameter("shisetuKbn", NpgsqlTypes.NpgsqlDbType.Varchar)) '施設区分
            If dataEXTM0102.PropTheaterBtn.Checked = True Then
                Adapter.SelectCommand.Parameters("shisetuKbn").Value = "1"
            Else
                Adapter.SelectCommand.Parameters("shisetuKbn").Value = "2"
            End If
            '期間from
            Adapter.SelectCommand.Parameters.Add(New NpgsqlParameter("bunruiCd", NpgsqlTypes.NpgsqlDbType.Varchar))                '分類コード
            Adapter.SelectCommand.Parameters("bunruiCd").Value = Format(Integer.Parse(dataEXTM0102.PropMaxBunruiCd), "00")         '分類コード

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
            puErrMsg = EXTM0102_E0000 & ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' シアター利用料用データ存在チェック用SQL作成
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataHBKA0101">[IN]データクラス</param>
    ''' </summary> 
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>付帯設備分類マスタから、情報を取得し、値をセットする（今はセットしていない）
    ''' <para>作成情報：2015/08/11 yu.satoh 
    ''' </para></remarks>
    Public Function GetRiyoKamokuInf(ByRef Adapter As NpgsqlDataAdapter, ByVal Cn As NpgsqlConnection, ByVal dataEXTM0102 As DataEXTM0102, ByRef i As Integer) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数の宣言
        Dim strSQL As String

        Try

            'SQL文(SELECT)
            strSQL = strSelectKamoku
            'Where句作成

            'データアダプタに、SQLを設定
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            'バインド変数をセット
            Adapter.SelectCommand.Parameters.Add(New NpgsqlParameter("KAMOKU_CD", NpgsqlTypes.NpgsqlDbType.Varchar))                 ' 科目コード
            Adapter.SelectCommand.Parameters.Add(New NpgsqlParameter("SAIMOKU_CD", NpgsqlTypes.NpgsqlDbType.Varchar))                ' 細目コード 
            Adapter.SelectCommand.Parameters.Add(New NpgsqlParameter("UCHI_CD", NpgsqlTypes.NpgsqlDbType.Varchar))                   ' 内訳コード
            Adapter.SelectCommand.Parameters.Add(New NpgsqlParameter("SHOSAI_CD", NpgsqlTypes.NpgsqlDbType.Varchar))                 ' 詳細コード
            Adapter.SelectCommand.Parameters("KAMOKU_CD").Value = dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 4).Value     ' 科目コード
            Adapter.SelectCommand.Parameters("SAIMOKU_CD").Value = dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 5).Value    ' 細目コード
            Adapter.SelectCommand.Parameters("UCHI_CD").Value = dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 6).Value       ' 内訳コード
            Adapter.SelectCommand.Parameters("SHOSAI_CD").Value = dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 7).Value     ' 細目コード
            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
            puErrMsg = EXTM0102_E0000 & ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' シアター利用料用科目情報取得用SQL作成
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataHBKA0101">[IN]データクラス</param>
    ''' </summary> 
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>付帯設備分類マスタから、情報を取得し、値をセットする（今はセットしていない）
    ''' <para>作成情報：2015/08/11 yu.satoh 
    ''' </para></remarks>
    Public Function SqlKamokuinf(ByRef Adapter As NpgsqlDataAdapter, ByVal Cn As NpgsqlConnection, ByVal dataEXTM0102 As DataEXTM0102) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数の宣言
        Dim strSQL As String

        Try

            'SQL文(SELECT)
            strSQL = strGetKamoku
            'Where句作成

            'データアダプタに、SQLを設定
            Adapter.SelectCommand = New NpgsqlCommand(strSQL, Cn)

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter.SelectCommand)
            puErrMsg = EXTM0102_E0000 & ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' 付帯表インサートSQL作成
    ''' <param name="Adapter">[IN/OUT]NpgSqlDataAdapterクラス</param>
    ''' <param name="Cn">[IN]NpgSqlConnectionクラス</param>
    ''' <param name="dataHBKA0101">[IN]データクラス</param>
    ''' </summary> 
    ''' <returns>boolean エラーコード  true 正常終了  false	異常終了 </returns>
    ''' <remarks>画面の値を、付帯設備マスタに登録する。
    ''' <para>作成情報：2015.12.25 h.hagiwara  
    ''' </para></remarks>
    Public Function InsertRiyoryoInf(ByRef Adapter As NpgsqlCommand, ByRef Tsx As NpgsqlTransaction, _
                                ByVal Cn As NpgsqlConnection, ByVal dataEXTM0102 As DataEXTM0102)

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数の宣言
        Dim strSQL As String
        '日付の設定
        Dim dtFrom As String = dataEXTM0102.PropYearFrom.Text + "/" + dataEXTM0102.PropMonthFrom.Text + "/01"
        Dim dtToSub As DateTime = dataEXTM0102.PropYearTo.Text + "/" + dataEXTM0102.PropMonthTo.Text
        '期間ＴＯは、入力された値の月＋１、日付-1の値で最終日を設定できる
        Dim dtTo As String = dtToSub.AddMonths(1).AddDays(-1).ToString("yyyy/MM/dd")

        Try

            'SQL文(INSERT)
            strSQL = strInsertBunrui

            'データアダプタに、SQLを設定
            Adapter = New NpgsqlCommand(strSQL, Cn)

            'バインド変数に型と値をセット
            Adapter.Parameters.Add(New NpgsqlParameter("kikanFrom", NpgsqlTypes.NpgsqlDbType.Varchar))      '期間from
            Adapter.Parameters.Add(New NpgsqlParameter("kikanTo", NpgsqlTypes.NpgsqlDbType.Varchar))        '期間to
            Adapter.Parameters.Add(New NpgsqlParameter("shisetuKbn", NpgsqlTypes.NpgsqlDbType.Varchar))     '施設区分
            Adapter.Parameters.Add(New NpgsqlParameter("bunruiCd", NpgsqlTypes.NpgsqlDbType.Varchar))       '分類コード
            Adapter.Parameters.Add(New NpgsqlParameter("bunruiNm", NpgsqlTypes.NpgsqlDbType.Varchar))       '分類名
            Adapter.Parameters.Add(New NpgsqlParameter("shukeiGrp", NpgsqlTypes.NpgsqlDbType.Varchar))      '集計グループ
            Adapter.Parameters.Add(New NpgsqlParameter("notaxFlg", NpgsqlTypes.NpgsqlDbType.Varchar))       '税
            Adapter.Parameters.Add(New NpgsqlParameter("sort", NpgsqlTypes.NpgsqlDbType.Integer))           '並び順
            Adapter.Parameters.Add(New NpgsqlParameter("kamokuCd", NpgsqlTypes.NpgsqlDbType.Varchar))       '科目コード
            Adapter.Parameters.Add(New NpgsqlParameter("saimokuCd", NpgsqlTypes.NpgsqlDbType.Varchar))      '細目コード
            Adapter.Parameters.Add(New NpgsqlParameter("uchiCd", NpgsqlTypes.NpgsqlDbType.Varchar))         '内訳コード
            Adapter.Parameters.Add(New NpgsqlParameter("shosaiCd", NpgsqlTypes.NpgsqlDbType.Varchar))       '詳細コード
            Adapter.Parameters.Add(New NpgsqlParameter("karikamokuCd", NpgsqlTypes.NpgsqlDbType.Varchar))   '借方科目コード
            Adapter.Parameters.Add(New NpgsqlParameter("karisaimokuCd", NpgsqlTypes.NpgsqlDbType.Varchar))  '借方細目コード
            Adapter.Parameters.Add(New NpgsqlParameter("kariuchiCd", NpgsqlTypes.NpgsqlDbType.Varchar))     '借方内訳コード
            Adapter.Parameters.Add(New NpgsqlParameter("karishosaiCd", NpgsqlTypes.NpgsqlDbType.Varchar))   '借方詳細コード
            Adapter.Parameters.Add(New NpgsqlParameter("sts", NpgsqlTypes.NpgsqlDbType.Varchar))            '無効フラグ
            Adapter.Parameters.Add(New NpgsqlParameter("RIYORYO_FLG", NpgsqlTypes.NpgsqlDbType.Varchar))    ' 付帯利用料フラグ
            Adapter.Parameters.Add(New NpgsqlParameter("addDate", NpgsqlTypes.NpgsqlDbType.Timestamp))      '作成日付
            Adapter.Parameters.Add(New NpgsqlParameter("addUserCd", NpgsqlTypes.NpgsqlDbType.Varchar))      '作成者名
            Adapter.Parameters.Add(New NpgsqlParameter("upDate", NpgsqlTypes.NpgsqlDbType.Timestamp))       '更新日
            Adapter.Parameters.Add(New NpgsqlParameter("upUserCd", NpgsqlTypes.NpgsqlDbType.Varchar))       '更新者

            '値をセット
            Adapter.Parameters("kikanFrom").Value = dtFrom  '日付from
            Adapter.Parameters("kikanTo").Value = dtTo      '日付to
            '施設区分がシアターなら1、スタジオなら2を設定
            If dataEXTM0102.PropTheaterBtn.Checked = True Then
                Adapter.Parameters("shisetuKbn").Value = "1"
            Else
                Adapter.Parameters("shisetuKbn").Value = "2"
            End If
            'Adapter.Parameters("bunruiCd").Value = dataEXTM0102.PropVwGroupingSheet.ActiveSheet.Cells(i, 0).Value '分類コード
            Adapter.Parameters("bunruiCd").Value = "99"
            Adapter.Parameters("bunruiNm").Value = "【使用不可】シアター利用料用" '分類名
            Adapter.Parameters("shukeiGrp").Value = "000000" '集計グループ
            Adapter.Parameters("notaxFlg").Value = "0"
            Adapter.Parameters("sort").Value = 99   '並び順
            Adapter.Parameters("kamokuCd").Value = dataEXTM0102.PropDtCopyKamokuMst.Rows(0).Item(0)       '科目コード
            Adapter.Parameters("saimokuCd").Value = dataEXTM0102.PropDtCopyKamokuMst.Rows(0).Item(1)      '細目コード
            Adapter.Parameters("uchiCd").Value = dataEXTM0102.PropDtCopyKamokuMst.Rows(0).Item(2)         '内訳コード
            Adapter.Parameters("shosaiCd").Value = dataEXTM0102.PropDtCopyKamokuMst.Rows(0).Item(3)       '細目コード
            Adapter.Parameters("karikamokuCd").Value = dataEXTM0102.PropDtCopyKamokuMst.Rows(0).Item(4)   '借方科目コード
            Adapter.Parameters("karisaimokuCd").Value = dataEXTM0102.PropDtCopyKamokuMst.Rows(0).Item(5)  '借方細目コード
            Adapter.Parameters("kariuchiCd").Value = dataEXTM0102.PropDtCopyKamokuMst.Rows(0).Item(6)     '借方内訳コード
            Adapter.Parameters("karishosaiCd").Value = dataEXTM0102.PropDtCopyKamokuMst.Rows(0).Item(7)   '借方詳細コード
            Adapter.Parameters("sts").Value = "0"
            Adapter.Parameters("RIYORYO_FLG").Value = "0"
            Adapter.Parameters("addDate").Value = Now
            Adapter.Parameters("addUserCd").Value = CommonEXT.PropComStrUserId
            Adapter.Parameters("upDate").Value = Now
            Adapter.Parameters("upUserCd").Value = CommonEXT.PropComStrUserId

            'コマンド実行
            Adapter.ExecuteNonQuery()

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常終了
            Return True

        Catch ex As Exception
            '例外発生
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Adapter)
            puErrMsg = EXTM0102_E0000 & ex.Message
            Return False
        End Try

    End Function

End Class
