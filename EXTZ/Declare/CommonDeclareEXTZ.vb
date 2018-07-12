''' <summary>
''' EXTZプロジェクトの共通定数
''' </summary>
''' <remarks>
''' EXTZプロジェクト内で共通で使用する定数を定義する。
''' <para>作成情報：2015/08/28 h.endo</para>
''' <p>改訂情報:</p>
''' </remarks>
Module CommonDeclareEXTZ
    '***************************************************************************************
    '* 共通変数
    '***************************************************************************************
    Public puErrMsg As String   'エラーメッセージ

    '***************************************************************************************
    '* 共通定数
    '***************************************************************************************
    Public Const TITLE As String = "EXTZ"   'タイトル

    '予約一覧
    '＜スプレッドの列Index＞
    '---シアター
    Public Const SpreadIndex_Theatre_Check As Integer = 0       '選択チェックボックス
    Public Const SpreadIndex_Theatre_YoyakuNo As Integer = 1 '予約番号
    Public Const SpreadIndex_Theatre_RiyoDtFrom As Integer = 2 '利用開始日
    Public Const SpreadIndex_Theatre_RiyoDtTo As Integer = 3 '利用終了日
    Public Const SpreadIndex_Theatre_StartTime As Integer = 4 '開始時間
    Public Const SpreadIndex_Theatre_EndTime As Integer = 5 '終了時間
    Public Const SpreadIndex_Theatre_SaijiNm As Integer = 6 '催事名
    Public Const SpreadIndex_Theatre_RiyoType As Integer = 7 '利用形状
    Public Const SpreadIndex_Theatre_RiyoNm As Integer = 8 '利用者名
    Public Const SpreadIndex_Theatre_SekininNm As Integer = 9 '責任者名
    Public Const SpreadIndex_Theatre_ShoninNinzu As Integer = 10 '承認人数
    Public Const SpreadIndex_Theatre_RiyoKin As Integer = 11 '利用料
    Public Const SpreadIndex_Theatre_FInputSts As Integer = 12 '設備入力ステータス
    Public Const SpreadIndex_Theatre_FutaiSetubi As Integer = 13 '付帯設備
    Public Const SpreadIndex_Theatre_Button As Integer = 14 '確認・編集ボタン
    '---スタジオ
    Public Const SpreadIndex_Studio_Check As Integer = 0 '選択チェックボックス
    Public Const SpreadIndex_Studio_YoyakuNo As Integer = 1 '予約番号
    Public Const SpreadIndex_Studio_Studio As Integer = 2 'スタジオ
    Public Const SpreadIndex_Studio_RiyoDtFrom As Integer = 3 '利用開始日
    Public Const SpreadIndex_Studio_RiyoDtTo As Integer = 4 '利用終了日
    Public Const SpreadIndex_Studio_StartTime As Integer = 5 '開始時間
    Public Const SpreadIndex_Studio_EndTime As Integer = 6 '終了時間
    Public Const SpreadIndex_Studio_ShutsuenNm As Integer = 7 'アーティスト名
    Public Const SpreadIndex_Studio_RiyoNm As Integer = 8 '利用者名
    Public Const SpreadIndex_Studio_SekininNm As Integer = 9 '責任者名
    Public Const SpreadIndex_Studio_ShoninNinzu As Integer = 10 '承認人数
    Public Const SpreadIndex_Studio_RiyoKin As Integer = 11 '利用料
    Public Const SpreadIndex_Studio_FInputSts As Integer = 12 '設備入力ステータス
    Public Const SpreadIndex_Studio_FutaiSetubi As Integer = 13 '付帯設備
    Public Const SpreadIndex_Studio_Button As Integer = 14 '確認・編集ボタン

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
    Public Const SpreadSeikyuIndex_Theatre_YoyakuNo As Integer = 10         '予約番号 
    Public Const SpreadSeikyuIndex_Theatre_Button As Integer = 11           '確認・編集ボタン
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
    Public Const SpreadSeikyuIndex_Studio_YoyakuNo As Integer = 11          '予約番号
    Public Const SpreadSeikyuIndex_Studio_Button As Integer = 12            '確認・編集ボタン

    '日別売上一覧
    '＜スプレッドの列Index＞
    '---シアター日別売上
    Public Const SpreadDayUriageIndex_Theatre_YoyakuNo As Integer = 0           '予約番号
    Public Const SpreadDayUriageIndex_Theatre_RiyoDt As Integer = 1             '利用日
    Public Const SpreadDayUriageIndex_Theatre_SaijiNm As Integer = 2            '催事名
    Public Const SpreadDayUriageIndex_Theatre_RiyoNm As Integer = 3             '利用者名
    Public Const SpreadDayUriageIndex_Theatre_KashiKind As Integer = 4          '貸出種別
    Public Const SpreadDayUriageIndex_Theatre_RiyoType As Integer = 5           '利用形状
    Public Const SpreadDayUriageIndex_Theatre_SaijiBunrui As Integer = 6        '催事分類
    Public Const SpreadDayUriageIndex_Theatre_RiyoKin As Integer = 7            '利用料
    Public Const SpreadDayUriageIndex_Theatre_FutaiKin As Integer = 8           '付帯設備使用料
    Public Const SpreadDayUriageIndex_Theatre_FutaiChosei As Integer = 9        '付帯設備調整額 
    Public Const SpreadDayUriageIndex_Theatre_KihonUriageGokei As Integer = 10  '基本売上合計
    Public Const SpreadDayUriageIndex_Theatre_DrinkGenkin As Integer = 11       'ドリンク現金
    Public Const SpreadDayUriageIndex_Theatre_OneDrink As Integer = 12          'ワンドリンク 
    Public Const SpreadDayUriageIndex_Theatre_CoinLocker As Integer = 13        'コインロッカー
    Public Const SpreadDayUriageIndex_Theatre_ZatsuShunyu As Integer = 14       '雑収入
    Public Const SpreadDayUriageIndex_Theatre_Aki1 As Integer = 15              '空き1
    Public Const SpreadDayUriageIndex_Theatre_Aki2 As Integer = 16              '空き2
    Public Const SpreadDayUriageIndex_Theatre_Aki3 As Integer = 17              '空き3
    Public Const SpreadDayUriageIndex_Theatre_Aki4 As Integer = 18              '空き4
    Public Const SpreadDayUriageIndex_Theatre_SuikaGenkin As Integer = 19       'SUIKA現金
    Public Const SpreadDayUriageIndex_Theatre_SuikaCoinLocker As Integer = 20   'SUIKAコインロッカー
    Public Const SpreadDayUriageIndex_Theatre_Sonota As Integer = 21            'その他
    Public Const SpreadDayUriageIndex_Theatre_GenkinGokei As Integer = 22       '現金合計
    Public Const SpreadDayUriageIndex_Theatre_SoUriageGokei As Integer = 23     '総売上合計 
    '---シアター請求調整
    Public Const SpreadSeikyuChoseiIndex_Theatre_YoyakuNo As Integer = 0        '予約番号
    Public Const SpreadSeikyuChoseiIndex_Theatre_RiyoDt As Integer = 1          '利用日
    Public Const SpreadSeikyuChoseiIndex_Theatre_SaijiNm As Integer = 2         '催事名
    Public Const SpreadSeikyuChoseiIndex_Theatre_RiyoNm As Integer = 3          '利用者名
    Public Const SpreadSeikyuChoseiIndex_Theatre_KashiKind As Integer = 4       '貸出種別
    Public Const SpreadSeikyuChoseiIndex_Theatre_RiyoType As Integer = 5        '利用形状 
    Public Const SpreadSeikyuChoseiIndex_Theatre_SaijiBunrui As Integer = 6     '催事分類
    Public Const SpreadSeikyuChoseiIndex_Theatre_SeikyuIraiNo As Integer = 7    '請求依頼番号
    Public Const SpreadSeikyuChoseiIndex_Theatre_SeikyuNaiyo As Integer = 8     '請求内容
    Public Const SpreadSeikyuChoseiIndex_Theatre_ChoseiKin As Integer = 9       '調整額
    '---スタジオ日別売上
    Public Const SpreadDayUriageIndex_Studio_YoyakuNo As Integer = 0            '予約番号
    Public Const SpreadDayUriageIndex_Studio_RiyoDt As Integer = 1              '利用日
    Public Const SpreadDayUriageIndex_Studio_ShutsuenNm As Integer = 2          'アーティスト名
    Public Const SpreadDayUriageIndex_Studio_RiyoNm As Integer = 3              '利用者名
    Public Const SpreadDayUriageIndex_Studio_KashiKind As Integer = 4           '貸出種別
    Public Const SpreadDayUriageIndex_Studio_Studio As Integer = 5              'スタジオ
    Public Const SpreadDayUriageIndex_Studio_RiyoKin As Integer = 6             '利用料
    Public Const SpreadDayUriageIndex_Studio_FutaiKin As Integer = 7            '付帯設備使用料
    Public Const SpreadDayUriageIndex_Studio_FutaiChosei As Integer = 8         '付帯設備調整額 
    Public Const SpreadDayUriageIndex_Studio_KihonUriageGokei As Integer = 9    '基本売上合計
    Public Const SpreadDayUriageIndex_Studio_DrinkGenkin As Integer = 10        'ドリンク現金
    Public Const SpreadDayUriageIndex_Studio_OneDrink As Integer = 11           'ワンドリンク 
    Public Const SpreadDayUriageIndex_Studio_CoinLocker As Integer = 12         'コインロッカー
    Public Const SpreadDayUriageIndex_Studio_ZatsuShunyu As Integer = 13        '雑収入
    Public Const SpreadDayUriageIndex_Studio_Aki1 As Integer = 14               '空き1
    Public Const SpreadDayUriageIndex_Studio_Aki2 As Integer = 15               '空き2
    Public Const SpreadDayUriageIndex_Studio_Aki3 As Integer = 16               '空き3
    Public Const SpreadDayUriageIndex_Studio_Aki4 As Integer = 17               '空き4
    Public Const SpreadDayUriageIndex_Studio_SuikaGenkin As Integer = 18        'SUIKA現金
    Public Const SpreadDayUriageIndex_Studio_SuikaCoinLocker As Integer = 19    'SUIKAコインロッカー
    Public Const SpreadDayUriageIndex_Studio_Sonota As Integer = 20             'その他
    Public Const SpreadDayUriageIndex_Studio_GenkinGokei As Integer = 21        '現金合計
    Public Const SpreadDayUriageIndex_Studio_SoUriageGokei As Integer = 22      '総売上合計 
    '---スタジオ請求調整
    Public Const SpreadSeikyuChoseiIndex_Studio_YoyakuNo As Integer = 0         '予約番号
    Public Const SpreadSeikyuChoseiIndex_Studio_RiyoDt As Integer = 1           '利用日
    Public Const SpreadSeikyuChoseiIndex_Studio_ShutsuenNm As Integer = 2       'アーティスト名
    Public Const SpreadSeikyuChoseiIndex_Studio_RiyoNm As Integer = 3           '利用者名
    Public Const SpreadSeikyuChoseiIndex_Studio_KashiKind As Integer = 4        '貸出種別
    Public Const SpreadSeikyuChoseiIndex_Studio_Studio As Integer = 5           'スタジオ 
    Public Const SpreadSeikyuChoseiIndex_Studio_SeikyuIraiNo As Integer = 6     '請求依頼番号
    Public Const SpreadSeikyuChoseiIndex_Studio_SeikyuNaiyo As Integer = 7      '請求内容
    Public Const SpreadSeikyuChoseiIndex_Studio_ChoseiKin As Integer = 8        '調整額

    '利用状況一覧
    '＜スプレッドの行Index＞
    '---シアター
    Public Const SpreadDayRiyoJokyoIndex_Theatre_RIYO_KANO_NISSU As Integer = 0             '利用可能日数
    Public Const SpreadDayRiyoJokyoIndex_Theatre_KYUKAN_NISSU As Integer = 1                '休館日数
    Public Const SpreadDayRiyoJokyoIndex_Theatre_MAINTE_NISSU As Integer = 2                'メンテナンス日数
    Public Const SpreadDayRiyoJokyoIndex_Theatre_RIYO_NISSU_TOTAL As Integer = 3            '使用日数(TOTAL)
    Public Const SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_TOTAL As Integer = 4             '利用率（TOTAL）
    Public Const SpreadDayRiyoJokyoIndex_Theatre_RIYO_NISSU_CHAKUSEKI As Integer = 5        '使用日数（着席）
    Public Const SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_CHAKUSEKI As Integer = 6         '利用率（着席）
    Public Const SpreadDayRiyoJokyoIndex_Theatre_RIYO_NISSU_STANDING As Integer = 7         '使用日数（スタンディング）
    Public Const SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_STANDING As Integer = 8          '利用率（スタンディング）
    Public Const SpreadDayRiyoJokyoIndex_Theatre_RIYO_NISSU_MIX As Integer = 9              '使用日数（MIX） 
    Public Const SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_MIX As Integer = 10              '利用率（MIX）
    Public Const SpreadDayRiyoJokyoIndex_Theatre_RIYO_NISSU_SAIJI As Integer = 11           '使用日数（催事） 
    Public Const SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_SAIJI As Integer = 12            '利用率（催事）
    Public Const SpreadDayRiyoJokyoIndex_Theatre_RIYO_NISSU_MUSIC As Integer = 13           '使用日数（音楽） 
    Public Const SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_MUSIC As Integer = 14            '利用率（音楽）
    Public Const SpreadDayRiyoJokyoIndex_Theatre_RIYO_NISSU_ENGEKI As Integer = 15          '使用日数（演劇）
    Public Const SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_ENGEKI As Integer = 16           '利用率（演劇）
    Public Const SpreadDayRiyoJokyoIndex_Theatre_RIYO_NISSU_ENGEI As Integer = 17           '使用日数（演芸）
    Public Const SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_ENGEI As Integer = 18            '利用率（演芸）
    Public Const SpreadDayRiyoJokyoIndex_Theatre_RIYO_NISSU_BUSINESS As Integer = 19        '使用日数（ビジネス）
    Public Const SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_BUSINESS As Integer = 20         '利用率（ビジネス）
    Public Const SpreadDayRiyoJokyoIndex_Theatre_RIYO_NISSU_MOVIE As Integer = 21           '使用日数（試写会・映画）
    Public Const SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_MOVIE As Integer = 22            '利用率（試写会・映画）
    Public Const SpreadDayRiyoJokyoIndex_Theatre_RIYO_NISSU_SONOTA As Integer = 23          '使用日数（その他）
    Public Const SpreadDayRiyoJokyoIndex_Theatre_RIYO_RITU_SONOTA As Integer = 24           '利用率（その他）
    '---スタジオ
    Public Const SpreadDayRiyoJokyoIndex_Studio_RIYO_KANO_NISSU As Integer = 0              '利用可能日数
    Public Const SpreadDayRiyoJokyoIndex_Studio_KYUKAN_NISSU As Integer = 1                 '休館日数
    Public Const SpreadDayRiyoJokyoIndex_Studio_MAINTE_NISSU As Integer = 2                 'メンテナンス日数
    Public Const SpreadDayRiyoJokyoIndex_Studio_RIYO_NISSU_TOTAL As Integer = 3             '使用日数(TOTAL)
    Public Const SpreadDayRiyoJokyoIndex_Studio_RIYO_RITU_TOTAL As Integer = 4              '利用率（TOTAL）
    Public Const SpreadDayRiyoJokyoIndex_Studio_RIYO_NISSU_201ST As Integer = 5             '使用日数（201st）
    Public Const SpreadDayRiyoJokyoIndex_Studio_RIYO_RITU_201ST As Integer = 6              '利用率（201st）
    Public Const SpreadDayRiyoJokyoIndex_Studio_RIYO_NISSU_202ST As Integer = 7             '使用日数（202st）
    Public Const SpreadDayRiyoJokyoIndex_Studio_RIYO_RITU_202ST As Integer = 8              '利用率（202st）
    Public Const SpreadDayRiyoJokyoIndex_Studio_RIYO_NISSU_HOUSELOCK As Integer = 9         '使用日数（house lock） 
    Public Const SpreadDayRiyoJokyoIndex_Studio_RIYO_RITU_HOUSELOCK As Integer = 10         '利用率（house lock）

End Module
