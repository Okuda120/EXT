Public Module CommonDeclareEXTY
    'EXAS請求依頼スプレッドシート（シアター）
    Public Const THEATER As Integer = 0
    'EXAS請求依頼スプレッドシート（スタジオ）
    Public Const STUDIO As Integer = 1

    'チェックボックス未選択エラーメッセージ
    Public Const Y0101_E0014 As String = "{0}を一つ以上選択してください。"

    'CSV出力時確認メッセージ
    Public Const Y0101_C0009 As String = "ファイル出力を行います。よろしいですか？"

    '処理完了時メッセージ
    Public Const Y0101_I0002 As String = "{0}が完了しました。"

    'CSV出力先フォルダ
    Public Const CSV_OUTPUT_FOLDER As String = "CSV/{0}"

    'ALSOK現金入金機データ登録画面（EXTY0102）
    Public Const Y0102_E001 As String = "レジマスタデータが取得できませんでした。"
    Public Const Y0102_E002 As String = "店舗マスタデータが取得できませんでした。"
    Public Const Y0102_E003 As String = "取り込みする入金ファイルが指定されていません。"
    Public Const Y0102_E004 As String = "指定された取り込みする入金ファイルが存在しません。"
    Public Const Y0102_E005 As String = "取り込み可能なファイルの拡張子は、ＣＳＶ形式のみです。"
    Public Const Y0102_E006 As String = "現金機利用日（ＦＲＯＭ）または現金機利用日（ＴＯ）を入力してください。"
    Public Const Y0102_E007 As String = "催事名／アーティスト名を設定してください。"
    Public Const Y0102_E008 As String = "更新する入金データが存在しません。"
    Public Const Y0102_E009 As String = "取り込み指定されたファイルの項目数が不正です。"
    Public Const Y0102_E010 As String = "現金機利用日（ＦＲＯＭ）には現金機利用日（ＴＯ）以降の日付を指定してください。"
    Public Const Y0102_E011 As String = "電子マネー入力で店舗が選択されていません。"
    Public Const Y0102_E012 As String = "電子マネー入力で投入金額が入力されていません。"

    Public Const Y0102_I001 As String = "入金データの登録が完了しました。"
    Public Const Y0102_I002 As String = "取り込みしたデータで{0}件が重複していました。"
    Public Const Y0102_I003 As String = "入金データ登録を行います。" & vbCrLf & "よろしいですか。"

    ' スプレッドシートの列（ALSOK入金情報表）
    Public Const COL_ALSOK_NYUKIN_DATE As Integer = 0          ' 入金日
    Public Const COL_ALSOK_SAIJI_NAME As Integer = 1           ' 催事名/アーティスト名
    Public Const COL_ALSOK_SELECT_BTN As Integer = 2           ' 選択ボタン
    Public Const COL_ALSOK_REGISTER_NO As Integer = 3          ' レジ№
    Public Const COL_ALSOK_REGISTER_NAME As Integer = 4        ' レジ名
    Public Const COL_ALSOK_TENPO_NO As Integer = 5             ' 店舗№
    Public Const COL_ALSOK_TENPO_NAME As Integer = 6           ' 店舗名
    Public Const COL_ALSOK_NYUKIN_GAKU As Integer = 7          ' 入金額
    Public Const COL_ALSOK_DAY_GASSAN As Integer = 8           ' 日毎合算
    Public Const COL_ALSOK_DELETE_BTN As Integer = 9           ' 削除ボタン
    Public Const COL_ALSOK_DEPOSIT_NO As Integer = 10          ' 入金機№
    Public Const COL_ALSOK_DEPOSIT_SEQ As Integer = 11         ' 入金機№連番
    Public Const COL_ALSOK_SHISETU_KBN As Integer = 12         ' 施設区分
    Public Const COL_ALSOK_PROP_KBN As Integer = 13            ' 削除区分
    Public Const COL_ALSOK_KEIYAKUSAKI As Integer = 14         ' 契約先

    ' スプレッドシートの列（電子マネー入金情報表）
    Public Const COL_CASH_NYUKIN_DATE As Integer = 0          ' 入金日
    Public Const COL_CASH_SAIJI_NAME As Integer = 1           ' 催事名/アーティスト名
    Public Const COL_CASH_SELECT_BTN As Integer = 2           ' 選択ボタン
    Public Const COL_CASH_REGISTER_NO As Integer = 3          ' レジ№
    Public Const COL_CASH_REGISTER_NAME As Integer = 4        ' レジ名
    Public Const COL_CASH_TENPO_NO As Integer = 5             ' 店舗№
    Public Const COL_CASH_TENPO_NAME As Integer = 6           ' 店舗名
    Public Const COL_CASH_NYUKIN_GAKU As Integer = 7          ' 入金額
    Public Const COL_CASH_DELETE_BTN As Integer = 8           ' 削除ボタン
    Public Const COL_CASH_DEPOSIT_NO As Integer = 9           ' 入金機№
    Public Const COL_CASH_DEPOSIT_SEQ As Integer = 10         ' 入金機№連番
    Public Const COL_CASH_SHISETU_KBN As Integer = 11         ' 施設区分
    Public Const COL_CASH_PROP_KBN As Integer = 12            ' 削除区分

    ' 電子マネー用入金機コード
    Public Const DEPOSITCASH_MACHINE_CD As String = "000000"  ' ALSOK入金情報表登録時の入金機コード
    Public Const CON_Y0102_CSV_ITEMCNT As Integer = 7         ' CSV項目数
    Public Const CON_Y0102_CSV1 As String = ".CSV"            ' 拡張子１
    Public Const CON_Y0102_CSV2 As String = ".csv"            ' 拡張子２

    '予約一覧
    '＜スプレッドの列Index＞
    '---シアター
    Public Const SpreadIndex_Theatre_Check As Integer = 0       '選択チェックボックス
    Public Const SpreadIndex_Theatre_Sts As Integer = 1 'ｽﾃｰﾀｽ
    Public Const SpreadIndex_Theatre_YoyakuNo As Integer = 2 '予約番号
    Public Const SpreadIndex_Theatre_RiyoDtFrom As Integer = 3 '利用開始日
    Public Const SpreadIndex_Theatre_RiyoDtTo As Integer = 4 '利用終了日
    Public Const SpreadIndex_Theatre_StartTime As Integer = 5 '開始時間
    Public Const SpreadIndex_Theatre_EndTime As Integer = 6 '終了時間
    Public Const SpreadIndex_Theatre_SaijiNm As Integer = 7 '催事名
    Public Const SpreadIndex_Theatre_RiyoType As Integer = 8 '利用形状
    Public Const SpreadIndex_Theatre_RiyoNm As Integer = 9 '利用者名
    Public Const SpreadIndex_Theatre_SekininNm As Integer = 10 '責任者名
    Public Const SpreadIndex_Theatre_ShoninNinzu As Integer = 11 '承認人数
    Public Const SpreadIndex_Theatre_RiyoKin As Integer = 12 '利用料
    Public Const SpreadIndex_Theatre_FInputSts As Integer = 13 '設備入力ステータス
    Public Const SpreadIndex_Theatre_FutaiSetubi As Integer = 14 '付帯設備
    Public Const SpreadIndex_Theatre_Button As Integer = 15 '確認・編集ボタン
    Public Const SpreadIndex_Theatre_StsCd As Integer = 16 'STS
    '---スタジオ
    Public Const SpreadIndex_Studio_Check As Integer = 0 '選択チェックボックス
    Public Const SpreadIndex_Studio_Sts As Integer = 1 'ｽﾃｰﾀｽ
    Public Const SpreadIndex_Studio_YoyakuNo As Integer = 2 '予約番号
    Public Const SpreadIndex_Studio_Studio As Integer = 3 'スタジオ
    Public Const SpreadIndex_Studio_RiyoDtFrom As Integer = 4 '利用開始日
    Public Const SpreadIndex_Studio_RiyoDtTo As Integer = 5 '利用終了日
    Public Const SpreadIndex_Studio_StartTime As Integer = 6 '開始時間
    Public Const SpreadIndex_Studio_EndTime As Integer = 7 '終了時間
    Public Const SpreadIndex_Studio_ShutsuenNm As Integer = 8 'アーティスト名
    Public Const SpreadIndex_Studio_RiyoNm As Integer = 9 '利用者名
    Public Const SpreadIndex_Studio_SekininNm As Integer = 10 '責任者名
    Public Const SpreadIndex_Studio_ShoninNinzu As Integer = 11 '承認人数
    Public Const SpreadIndex_Studio_RiyoKin As Integer = 12 '利用料
    Public Const SpreadIndex_Studio_FInputSts As Integer = 13 '設備入力ステータス
    Public Const SpreadIndex_Studio_FutaiSetubi As Integer = 14 '付帯設備
    Public Const SpreadIndex_Studio_Button As Integer = 15 '確認・編集ボタン
    Public Const SpreadIndex_Studio_StsCd As Integer = 16 'STS

    '請求一覧
    '＜スプレッドの列Index＞
    '---シアター
    Public Const SpreadSeikyuIndex_Theatre_SeikyuIraiNo As Integer = 0      '請求依頼番号
    Public Const SpreadSeikyuIndex_Theatre_SeikyuDt As Integer = 1          '請求日
    Public Const SpreadSeikyuIndex_Theatre_RiyoDtFrom As Integer = 2        '利用開始日
    Public Const SpreadSeikyuIndex_Theatre_RiyoDtTo As Integer = 3          '利用終了日
    Public Const SpreadSeikyuIndex_Theatre_SaijiNm As Integer = 4           '催事名
    Public Const SpreadSeikyuIndex_Theatre_AitesakiNm As Integer = 5        '相手先名
    Public Const SpreadSeikyuIndex_Theatre_SeikyuKingaku As Integer = 6     '請求金額 
    Public Const SpreadSeikyuIndex_Theatre_SeikyuNaiyo As Integer = 7       '請求内容
    Public Const SpreadSeikyuIndex_Theatre_NyukinYoteiDt As Integer = 8     '入金予定日
    Public Const SpreadSeikyuIndex_Theatre_NyukinDt As Integer = 9          '入金日
    ' 2016.01.13 ADD START↓ y.morooka 入金情報複数件の対応
    Public Const SpreadSeikyuIndex_Theatre_NyukinKin As Integer = 10        '入金額
    Public Const SpreadSeikyuIndex_Theatre_NyukinSts As Integer = 11        '入金状況
    ' 2016.01.13 ADD END↑ y.morooka 入金情報複数件の対応
    Public Const SpreadSeikyuIndex_Theatre_YoyakuNo As Integer = 12         '予約番号 
    Public Const SpreadSeikyuIndex_Theatre_Button As Integer = 13           '確認・編集ボタン
    '---スタジオ
    Public Const SpreadSeikyuIndex_Studio_SeikyuIraiNo As Integer = 0       '請求依頼番号
    Public Const SpreadSeikyuIndex_Studio_SeikyuDt As Integer = 1           '請求日
    Public Const SpreadSeikyuIndex_Studio_RiyoDtFrom As Integer = 2         '利用開始日
    Public Const SpreadSeikyuIndex_Studio_RiyoDtTo As Integer = 3           '利用終了日
    Public Const SpreadSeikyuIndex_Studio_Studio As Integer = 4             'スタジオ
    Public Const SpreadSeikyuIndex_Studio_ShutsuenNm As Integer = 5         'アーティスト名
    Public Const SpreadSeikyuIndex_Studio_AitesakiNm As Integer = 6         '相手先名
    Public Const SpreadSeikyuIndex_Studio_SeikyuKingaku As Integer = 7      '請求金額 
    Public Const SpreadSeikyuIndex_Studio_SeikyuNaiyo As Integer = 8        '請求内容
    Public Const SpreadSeikyuIndex_Studio_NyukinYoteiDt As Integer = 9      '入金予定日
    Public Const SpreadSeikyuIndex_Studio_NyukinDt As Integer = 10          '入金日
    ' 2016.01.13 ADD START↓ y.morooka 入金情報複数件の対応
    Public Const SpreadSeikyuIndex_Studio_NyukinKin As Integer = 11        '入金額
    Public Const SpreadSeikyuIndex_Studio_NyukinSts As Integer = 12        '入金状況
    ' 2016.01.13 ADD END↑ y.morooka 入金情報複数件の対応
    Public Const SpreadSeikyuIndex_Studio_YoyakuNo As Integer = 13          '予約番号
    Public Const SpreadSeikyuIndex_Studio_Button As Integer = 14            '確認・編集ボタン

End Module

