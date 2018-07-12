''' <summary>
''' EXTBプロジェクトの共通定数
''' </summary>
''' <remarks>
''' EXTBプロジェクト内で共通で使用する定数を定義する。
''' <para>作成情報：2015/08/07 h.endo</para>
''' <p>改訂情報:</p>
''' </remarks>
Module CommonDeclareEXTB
    '***************************************************************************************
    '* 共通変数
    '***************************************************************************************
    Public puErrMsg As String   'エラーメッセージ

    '***************************************************************************************
    '* 共通定数
    '***************************************************************************************
    Public Const TITLE As String = "EXTB"   'タイトル

    'キャンセルのモード
    Public Const CANCEL_WAITALL = "1"                   'キャンセル待ち（全て）
    Public Const CANCEL_WAITDATEONLY = "2"              'キャンセル待ち（日付確定のみ、年月）
    Public Const CANCEL_WAITDATEONLY_NENGAPPI = "3"     'キャンセル待ち（日付確定のみ、年月日）
    Public Const CANCEL_SUMI = "4"                      'キャンセル済

    '利用日検索モード
    Public Const YOYAKU_NENGAPPI = "1"      '年月日
    Public Const YOYAKU_NENGETU = "2"       '年月

    '休館メンテ日検索モード
    Public Const KYUKANMAINTE_NENGAPPI = "1"    '年月日
    Public Const KYUKANMAINTE_NENGETU = "2"     '年月

    '日付設定モード
    Public Const SETDATE_KYUKAN = "休館日"         '休館日
    Public Const SETDATE_EIGYO = "営業日"          '営業日
    Public Const SETDATE_MAINTE = "メンテ日"       'メンテ日

    '祝日設定モード
    Public Const HOLIYDAY_OFF = "平日"            '平日設定
    Public Const HOLIYDAY_ON = "祝日（休日）"     '祝日設定

    ' 2015.12.15 UPD START↓ h.hagiwara システム名部分変更
    'Public Const MAIL_SHONIN_IRAI_SUBJECT As String = "【施設予約管理システム】承認依頼通知"
    'Public Const MAIL_SHONIN_IRAI_BODY As String = _
    '                                        "以下の予約にて承認依頼が来ておりますので" & vbCrLf & _
    '                                        "お知らせいたします。" & vbCrLf & _
    '                                        "" & vbCrLf & _
    '                                        "－－－－－－－－－－－－－－－－－－－－－－－－" & vbCrLf & _
    '                                        "予約番号　：　［{0}］" & vbCrLf & _
    '                                        "催事名　：　［{1}］" & vbCrLf & _
    '                                        "利用者名　：　［{2}］" & vbCrLf & _
    '                                        "利用日　：　［{3}］～［{4}］" & vbCrLf & _
    '                                        "－－－－－－－－－－－－－－－－－－－－－－－－" & vbCrLf & _
    '                                        "" & vbCrLf & _
    '                                        "" & vbCrLf & _
    '                                        "本予約内容を確認する手順は以下の通りです。" & vbCrLf & _
    '                                        "１．施設予約管理システムにログイン" & vbCrLf & _
    '                                        "２．カレンダーから該当する予約を探し、ダブルクリック、" & vbCrLf & _
    '                                        "　　予約内容を閲覧する" & vbCrLf & _
    '                                        "　　または、" & vbCrLf & _
    '                                        "　　予約一覧で該当する予約を検索、該当する予約をダブルクリックし、" & vbCrLf & _
    '                                        "　　予約内容を閲覧する" & vbCrLf & _
    '                                        "３．予約内容を確認し、承認タブにて承認／差し戻しをする"


    'Public Const MAIL_SHONIN_KAKUNIN_SUBJECT As String = "【施設予約管理システム】承認／差し戻し結果通知"
    'Public Const MAIL_SHONIN_KAKUNIN_BODY As String = _
    '                                        "以下の予約にて承認／差し戻しが行われましたので" & vbCrLf & _
    '                                        "お知らせいたします。" & vbCrLf & _
    '                                        "" & vbCrLf & _
    '                                        "－－－－－－－－－－－－－－－－－－－－－－－－" & vbCrLf & _
    '                                        "予約番号　：　［{0}］" & vbCrLf & _
    '                                        "催事名　：　［{1}］" & vbCrLf & _
    '                                        "利用者名　：　［{2}］" & vbCrLf & _
    '                                        "利用日　：　［{3}］～［{4}］" & vbCrLf & _
    '                                        "－－－－－－－－－－－－－－－－－－－－－－－－" & vbCrLf & _
    '                                        "" & vbCrLf & _
    '                                        "" & vbCrLf & _
    '                                        "本予約内容を確認する手順は以下の通りです。" & vbCrLf & _
    '                                        "１．施設予約管理システムにログイン" & vbCrLf & _
    '                                        "２．カレンダーから該当する予約を探し、ダブルクリック、" & vbCrLf & _
    '                                        "　　予約内容の承認タブにて承認／差し戻し結果を確認する" & vbCrLf & _
    '                                        "　　または、" & vbCrLf & _
    '                                        "　　予約一覧で該当する予約を検索、該当する予約をダブルクリックし、" & vbCrLf & _
    '                                        "　　予約内容の承認タブにて承認／差し戻し結果を確認する"
    Public Const MAIL_SHONIN_IRAI_SUBJECT As String = "【ちゃり～ん。】承認依頼通知"
    Public Const MAIL_SHONIN_IRAI_BODY As String = _
                                            "以下の予約にて承認依頼が来ておりますので" & vbCrLf & _
                                            "お知らせいたします。" & vbCrLf & _
                                            "" & vbCrLf & _
                                            "－－－－－－－－－－－－－－－－－－－－－－－－" & vbCrLf & _
                                            "予約番号　：　［{0}］" & vbCrLf & _
                                            "催事名　　：　［{1}］" & vbCrLf & _
                                            "利用者名　：　［{2}］" & vbCrLf & _
                                            "利用日　　：　［{3}］～［{4}］" & vbCrLf & _
                                            "－－－－－－－－－－－－－－－－－－－－－－－－" & vbCrLf & _
                                            "" & vbCrLf & _
                                            "" & vbCrLf & _
                                            "本予約内容を確認する手順は以下の通りです。" & vbCrLf & _
                                            "１．【ちゃり～ん。】にログイン" & vbCrLf & _
                                            "２．カレンダーから該当する予約を探し、ダブルクリック、" & vbCrLf & _
                                            "　　予約内容を閲覧する" & vbCrLf & _
                                            "　　または、" & vbCrLf & _
                                            "　　予約一覧で該当する予約を検索、該当する予約の「確認・編集」ボタンをクリックし、" & vbCrLf & _
                                            "　　予約内容を閲覧する" & vbCrLf & _
                                            "３．予約内容を確認し、承認タブにて承認／差し戻しをする"


    Public Const MAIL_SHONIN_KAKUNIN_SUBJECT As String = "【ちゃり～ん。】承認／差し戻し結果通知"
    Public Const MAIL_SHONIN_KAKUNIN_BODY As String = _
                                            "以下の予約にて承認／差し戻しが行われましたので" & vbCrLf & _
                                            "お知らせいたします。" & vbCrLf & _
                                            "" & vbCrLf & _
                                            "－－－－－－－－－－－－－－－－－－－－－－－－" & vbCrLf & _
                                            "予約番号　：　［{0}］" & vbCrLf & _
                                            "催事名　　：　［{1}］" & vbCrLf & _
                                            "利用者名　：　［{2}］" & vbCrLf & _
                                            "利用日　　：　［{3}］～［{4}］" & vbCrLf & _
                                            "－－－－－－－－－－－－－－－－－－－－－－－－" & vbCrLf & _
                                            "" & vbCrLf & _
                                            "" & vbCrLf & _
                                            "本予約内容を確認する手順は以下の通りです。" & vbCrLf & _
                                            "１．【ちゃり～ん。】にログイン" & vbCrLf & _
                                            "２．カレンダーから該当する予約を探し、ダブルクリック、" & vbCrLf & _
                                            "　　予約内容の承認タブにて承認／差し戻し結果を確認する" & vbCrLf & _
                                            "　　または、" & vbCrLf & _
                                            "　　予約一覧で該当する予約を検索、該当する予約の「確認・編集」ボタンをクリックし、" & vbCrLf & _
                                            "　　予約内容の承認タブにて承認／差し戻し結果を確認する"
    ' 2015.12.15 UPD END↑ h.hagiwara システム名部分変更

    Public Const EXT_C0008 As String = "印刷します。よろしいですか？"
    '出力メッセージ
    Public Const EXT_C0009 As String = "ファイルを出力します。よろしいですか？"

    Public Const EXT_I0002 As String = "{0}が完了しました。"

    'CSVファイルヘッダ行文字列
    Public Const CSV_HEADER As String = "請求日,入金予定日,当社担当者コード,当社担当者所属部署コード,当社部署," &
                                         "担当者TEL,請求先コード,請求先郵便番号,請求先住所１,請求先住所２," &
                                         "請求先名称,G請求先部署コード,請求先部署名,G請求書担当者コード,請求書担当者名," &
                                         "経理連絡欄,タイトル１,タイトル２,納品書備考,処理区分," &
                                         "番組コード,番組シーケンス,プロジェクトコード,プロジェクトシーケンス,目的コード," &
                                         "預り先コード,計上部署コード,識別区分,コンテンツコード,コンテンツ内訳コード," &
                                         "予算外案件コード,勘定科目コード,細目コード,内訳コード,詳細コード," &
                                         "借方勘定科目コード,借方細目コード,借方内訳コード,借方詳細コード,発生月自," &
                                         "発生月至,セグメントコード,入力摘要１,入力摘要２,単価," &
                                         "数量,消費税額,消費税区分,消費税率,外税内税区分," &
                                         "G請求内容コード,Gセグメントコード,Gコンテンツ識別区分,Gコンテンツコード,Gコンテンツ内訳コード"

End Module
