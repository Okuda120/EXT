''' <summary>
''' EXTプロジェクトの共通定数
''' </summary>
''' <remarks>
''' EXTプロジェクト内で共通で使用する定数を定義する。
''' <para>作成情報：2015/08/04 k.machida</para>
''' <p>改訂情報:</p>
''' </remarks>
Public Module CommonDeclareEXT
    '***************************************************************************************
    '* 共通定数
    '***************************************************************************************
    Public Const TITLE As String = "EXT"   'タイトル
    Public Const TITLE_INFO As String = "メッセージ"
    Public Const TITLE_WARNING As String = "警告"
    Public Const TITLE_ERROR As String = "エラー"

    '***************************************************************************************
    '* フォーム背景色
    '***************************************************************************************
    Private FORM_BACKCOLOR_HONBAN As Color = Color.White                       ' 本番機
    Private FORM_BACKCOLOR_KENSHOU As Color = Color.FromArgb(192, 255, 255)    ' 検証機
    Private ppConfigrationFlg As String                                        ' 環境設定フラグ

    '***************************************************************************************
    '* Loginプロパティ
    '***************************************************************************************
    Private ppStrComUserId As String                   'ユーザーID
    Private ppStrComUserName As String                 'ユーザー名
    Private ppStrComMailAddr As String                 'メール
    Private ppStrComBusho As String                    '部署CD
    Private ppStrComFlgShonin As String                '承認FLG
    Private ppStrComFlgMst As String                   'マスタ操作FLG
    '***************************************************************************************
    '* Systemプロパティ
    '***************************************************************************************
    Private ppIntTeiinA As Integer                   'シアター定員
    Private ppIntTeiinB As Integer                   'シアター定員
    Private ppIntKariJizenTuti As Integer           '仮予約事前通知日
    Private ppGSeikyusaki As String                'グループ請求先        ' 2016.01.18 ADD y.morooka グループ請求対応
    Private ppStrDaikoFlg As String                '代行フラグ            ' 2016.08.12 ADD m.hayabuchi 代行処理対応
    Private ppStrDaikoTanto As String              '代行担当            ' 2016.09.20 ADD m.hayabuchi 代行処理対応
    Private ppStrDaikoBusho As String              '代行部署            ' 2016.09.20 ADD m.hayabuchi 代行処理対応
    '***************************************************************************************
    '* DBレプリケーションプロパティ
    '***************************************************************************************
    Private ppMainDB As String
    Private ppDummyFilePath1 As String
    Private ppDummyFilePath2 As String
    Private ppDummyFileName As String
    Private ppDBString1 As String
    Private ppDBString2 As String
    Private ppNetUseUserID1 As String
    Private ppNetUsePassword1 As String
    Private ppNetUseUserID2 As String
    Private ppNetUsePassword2 As String
    Private ppDBCheck As String
    Private ppDebugDB As String

    '*--------------------------------
    '* 定数
    '*--------------------------------
    '施設区分
    Public Const SHISETU_KBN_THEATER As String = "1"
    Public Const SHISETU_KBN_STUDIO As String = "2"
    Public Const SHISETU_KBN_BLANK As String = "0"
    'STUDIO区分
    Public Const STUDIO_MITEI As String = "0"
    Public Const STUDIO_201 As String = "1"
    Public Const STUDIO_202 As String = "2"
    Public Const STUDIO_HOUSE_LOCK As String = "3"
    '1-仮予約（未確認） 2-仮予約（確認済み） 3-正式予約（進行中） 4-正式予約（完了） 5-仮予約ｷｬﾝｾﾙ 6-正式予約ｷｬﾝｾﾙ
    Public Const YOYAKU_STS_KARI_MI As String = "1"
    Public Const YOYAKU_STS_KARI As String = "2"
    Public Const YOYAKU_STS_SEISHIKI As String = "3"
    Public Const YOYAKU_STS_SEISHIKI_COMP As String = "4"
    Public Const YOYAKU_STS_CANCEL_KARI As String = "5"
    Public Const YOYAKU_STS_CANCEL_SEISHIKI As String = "6"
    '1-ｷｬﾝｾﾙ待ち 2-（仮）予約昇格 3-ｷｬﾝｾﾙ取消
    Public Const CANCEL_STS_MATI As String = "1"
    Public Const CANCEL_STS_KARI As String = "2"
    Public Const CANCEL_STS_TORIKESI As String = "3"
    '利用者レベル 1-一般 2-要注意 3-利用不可
    Public Const RIYOSHA_LV1 As String = "1"
    Public Const RIYOSHA_LV2 As String = "2"
    Public Const RIYOSHA_LV3 As String = "3"
    '日付フォーマット
    Public Const FMT_DATE As String = "yyyy/MM/dd"
    Public Const FMT_DATE_DISP As String = "yyyy/MM/dd(ddd)"
    Public Const FMT_DATE_YM As String = "yyyy/MM"
    '利用形態 1-時間貸し 2-lock out
    Public Const RIYOKEITAI_TIME = "1"
    Public Const RIYOKEITAI_LOCKOUT = "2"
    Public Const RIYOKEITAI_TIME_NM = "時間貸し"
    Public Const RIYOKEITAI_LOCKOUT_NM = "Lock out"
    '休館メンテ区分 1-休館日 2-メンテナンス
    Public Const HOLMENT_KBN_HOLIDAY = "1"
    Public Const HOLMENT_KBN_MENTDAY = "2"
    '予約一覧の遷移元情報
    '---選択モード
    Public Const SELECTMODE_MENU As String = "0"         'メニュー画面からの遷移
    Public Const SELECTMODE_ALSOKNYUKIN As String = "1"  'ALSOK現金入金機データ登録画面からの遷移
    '*--------------------------------
    '* 定数:予約
    '*--------------------------------
    '利用日追加画面：登録区分
    Public Const TOUROKU_KBN_KARI As String = "0"
    Public Const TOUROKU_KBN_SEISHIKI As String = "1"
    Public Const TOUROKU_KBN_CANCEL As String = "2"
    '予約取消区分
    Public Const TORIKESHI_KBN_CANCEL As String = "1"
    Public Const TORIKESHI_KBN_DELETE As String = "2"
    '貸出種別 0-未定 1-一般貸出 2-社内利用 3-特例
    Public Const KASHI_KIND_MITEI As String = "0"
    Public Const KASHI_KIND_IPPAN As String = "1"
    Public Const KASHI_KIND_HOUSE As String = "2"
    Public Const KASHI_KIND_TOKUREI As String = "3"
    'シアター利用形状 0-未定 1-A.イベント（スタンディング） 2-B.イベント（着席） 3-A+Bイベント（MIX） 4-C.催事
    Public Const RIYO_TYPE_MITEI As String = "0"
    Public Const RIYO_TYPE_STAND As String = "1"
    Public Const RIYO_TYPE_SEATING As String = "2"
    Public Const RIYO_TYPE_MIX As String = "3"
    Public Const RIYO_TYPE_SAIJI As String = "4"
    'ワンドリンク 0-未定 1-有 2-無
    Public Const ONE_DRINK_MITEI As String = "0"
    Public Const ONE_DRINK_ARI As String = "1"
    Public Const ONE_DRINK_NASHI As String = "2"
    'シアター催事分類 0-未定 1-音楽 2-演劇 3-演芸 4-ビジネス 5-試写会・映画 6-その他
    Public Const SAIJI_MITEI As String = "0"
    Public Const SAIJI_MUSIC As String = "1"
    Public Const SAIJI_ENGEKI As String = "2"
    Public Const SAIJI_ENGEI As String = "3"
    Public Const SAIJI_BUSINESS As String = "4"
    Public Const SAIJI_MOVIE As String = "5"
    Public Const SAIJI_ETC As String = "6"
    '音響オペレータ 0-未定 1-有 2-無
    Public Const ONKYO_OPE_MITEI As String = "0"
    Public Const ONKYO_OPE_ARI As String = "1"
    Public Const ONKYO_OPE_NASHI As String = "2"
    '申込書送付区分 0-未定 1-ダウンロード 2-送付希望
    Public Const SEND_KBN_MITEI As String = "0"
    Public Const SEND_KBN_DL As String = "1"
    Public Const SEND_KBN_KIBOU As String = "2"
    '送付ｽﾃｰﾀｽ 0-未定 1-未送付 2-送付済
    Public Const SEND_STS_MITEI As String = "0"
    Public Const SEND_STS_NO_SHIPPED As String = "1"
    Public Const SEND_STS_SHIPPING As String = "2"
    'チケット入場料 0-未定 1-有料 2-無料
    Public Const TICKET_ENTER_MITEI As String = "0"
    Public Const TICKET_ENTER_YURYO As String = "1"
    Public Const TICKET_ENTER_MURYO As String = "2"
    'チケットドリンク代 0-未定 1-込 2-別
    Public Const TICKET_DRINK_NASHI As String = "0"
    Public Const TICKET_DRINK_KOMI As String = "1"
    Public Const TICKET_DRINK_BETSU As String = "2"
    'HP掲載 0-未定 1-可 2-不可
    Public Const HP_MITEI As String = "0"
    Public Const HP_OK As String = "1"
    Public Const HP_NG As String = "2"
    '付属設備入力ｽﾃｰﾀｽ 仮予約時NULL　0-未確定 1-確定
    Public Const F_INPUT_MIKAKU As String = "0"
    Public Const F_INPUT_KAKUTEI As String = "1"
    '承認区分
    Public Const SHONIN_IRAI As String = "1"
    Public Const SHONIN_KIROKU As String = "2"
    '請求内容
    Public Const SEIKYU_NAIYOU_RIYO As String = "1"
    Public Const SEIKYU_NAIYOU_FUTAI As String = "2"
    Public Const SEIKYU_NAIYOU_RIYOFUTAI As String = "3"
    Public Const SEIKYU_NAIYOU_ETC As String = "4"
    '*--------------------------------
    '* 定数:DBレプリケーション
    '*--------------------------------
    'NetUseユーザID(1号機)
    Public Const NET_USE_USERID1 As String = "vms00242\etr_masterflag"
    'NetUseパスワード(1号機)
    Public Const NET_USE_PASSWORD1 As String = "etr_flag2015"
    'NetUseユーザID(2号機)
    Public Const NET_USE_USERID2 As String = "vms00243\etr_masterflag"
    'NetUseパスワード(2号機)
    Public Const NET_USE_PASSWORD2 As String = "etr_flag2015"
    '*--------------------------------
    '* 定数:操作ログ
    '*--------------------------------
    'NetUseユーザID
    Public Const OPLOG_USE_USERID As String = "vms00241\etr_capture"
    'NetUseパスワード
    Public Const OPLOG_USE_PASSWORD As String = "ETR_capture2015"
    '*--------------------------------
    '* 定数:キャプチャフォルダ表示
    '*--------------------------------
    'NetUseユーザID
    Public Const CAPTURE_USE_USERID As String = "vms00241\etr_capture"
    'NetUseパスワード
    Public Const CAPTURE_USE_PASSWORD As String = "ETR_capture2015"
    '*--------------------------------
    '* エラーメッセージ
    '*--------------------------------
    '種別:エラー
    Public Const E0000 As String = "システムエラーが発生しました。システム管理者に連絡してください。"
    Public Const E0001 As String = "{0}を入力してください。"
    Public Const E0002 As String = "{0}を選択してください。"
    Public Const E0003 As String = "{0}は半角数字で入力してください。"
    Public Const E0004 As String = "{0}は半角数字または記号で入力してください。"
    Public Const E0005 As String = "{0}は半角英数字で入力してください。"
    Public Const E0006 As String = "{0}は全角文字で入力してください。"
    Public Const E0007 As String = "{0}は全角カナで入力してください。"
    Public Const E0008 As String = "入力された{0}は不正なアドレスのようです。再度入力し直してください。"
    Public Const E0009 As String = "{0}は{1}桁以上{2}桁以内で入力してください。"
    Public Const E0010 As String = "{0}は{1}桁で入力してください。"
    Public Const E0011 As String = "{0}と{1}に入力された値が一致しません。"
    Public Const E0012 As String = "{0}に入力された値が正しくありません。"
    Public Const E0013 As String = "{0}には過去の取得年を入力してください。"
    Public Const E0014 As String = "{0}を一つ以上選択してください。"
    Public Const E0015 As String = "画面の記載内容をご了承の上、{0}を選択してください。"
    Public Const E0016 As String = "{0}が受験可能な年齢の範囲外です。"
    Public Const E0017 As String = "{0}をアップロードしてください。"
    Public Const E0018 As String = "{0}は{1}から{2}の範囲内で入力してください。"
    Public Const E0019 As String = "{0}は{1}以下の値を入力してください。"
    Public Const E0020 As String = "{0}は{1}より小さい値を入力してください。"
    Public Const E0021 As String = "{0}は{1}以上の値を入力してください。"
    Public Const E0022 As String = "{0}は{1}より大きい値を入力してください。"
    Public Const E0023 As String = "{0}と{1}は同時に選択できません。"
    Public Const E0024 As String = "{0}は{1}桁以内で入力してください。"
    Public Const E0025 As String = "データ登録処理でエラーが発生しました。"
    Public Const E0026 As String = "{0}は正しいメールアドレスではありません。"

    Public Const E9999 As String = "データベースに接続できませんでした。"

    Public Const E1001 As String = "入力されたIDまたはパスワードが正しくありません。"
    Public Const E1002 As String = "入力された{0}は既に使用されています。"
    Public Const E1003 As String = "指定されたファイルの拡張子は、jpgまたはjpeg形式のみ有効です。"
    Public Const E1004 As String = "ファイルのサイズは50KBまでアップロード可能です。"
    Public Const E1005 As String = "入力された漢字氏名またはメールアドレスが正しくありません。"
    Public Const E1006 As String = "入力されたセキュリティ回答が正しくありません。"
    Public Const E1007 As String = "選択された受験日で、既に申込を行っています。"
    Public Const E1008 As String = "入力された受験番号が正しくありません。"
    Public Const E1009 As String = "入力された受験番号の受験日が現在日付を過ぎているため、印刷できません。"
    Public Const E1010 As String = "入力された受験番号の受験票が既に５回印刷されているため、印刷できません。"
    Public Const E2001 As String = "入力されたユーザIDまたはパスワードが正しくありません。"
    Public Const E2002 As String = "指定されたファイルが存在しません。"
    Public Const E2003 As String = "{0}の年月日は全て入力してください。"
    Public Const E2004 As String = "{0}の年月日時分は全て入力してください。"
    Public Const E2005 As String = "{0}の日付は正しくありません。"
    Public Const E2006 As String = "{0}の日付／時刻は正しくありません。"
    Public Const E2007 As String = "{0}の時刻は正しくありません。"
    Public Const E2008 As String = "{0}の時分は両方入力してください。"
    Public Const E2009 As String = "{0}を入力する場合、{1}も入力する必要があります。"
    Public Const E2010 As String = "{0}は{1}以前の日付を入力してください。"
    Public Const E2011 As String = "{0}は{1}以降の日付を入力してください。"
    Public Const E2012 As String = "指定されたファイルを読み込めません。"
    Public Const E2013 As String = "指定されたファイルにはデータ行が存在しません。"
    Public Const E2014 As String = "ファイルの拡張子は、csv形式のみ有効です。"
    Public Const E2015 As String = "利用日時は0件には出来ません。"
    Public Const E2016 As String = "既に利用者マスタに登録されています。マスタ検索から閲覧してください。"
    Public Const E2017 As String = "登録作業中の予約含め、最大予約数に達しています。指定日を予約登録できません。"
    Public Const E2018 As String = "利用者がマスタ登録されていません。利用者マスタに登録してから正式予約登録へ進んでください。"
    Public Const E2019 As String = "催事情報や日程に未定が残っているか、仮予約内部確認が未確認となっているため、正式予約に進めません。"
    Public Const E2020 As String = "{0}1～3の内、いずれかの番号が未入力です。"
    Public Const E2021 As String = "催事情報や日程に未定が残っているため、正式予約に進めません。"
    Public Const E2022 As String = "相手先が紐付いていないか、使用停止のため請求依頼登録できません。利用者マスタより紐付けを行って下さい。"
    Public Const E2023 As String = "選択した請求は未請求のため、処理を実行できません。"
    Public Const E2024 As String = "プロジェクトと紐付いていないため、完了できません。"
    Public Const E2025 As String = "承認書・利用明細・EXAS請求依頼ファイルの出力をおこなった場合は更新処理を行って下さい。"
    Public Const E2026 As String = "入力された年月は存在しません。"
    Public Const E2027 As String = "休館日またはメンテ日が選択されています。日付の選択を見直してください。"
    Public Const E2028 As String = "最大キャンセル待ち数に達しています。指定日のキャンセル待ちを登録できません。"
    Public Const E2029 As String = "{0}は1日ずつ指定してください。"
    Public Const E2030 As String = "既に仮予約または正式予約が登録されています。登録状況を確認してください。"
    Public Const E2031 As String = "既にこの日は{0}に指定されています。確認してください。"
    Public Const E2032 As String = "入力された期間は設定済みのため登録できません。"
    Public Const E2033 As String = "{0}があるため、入力完了できません。"
    Public Const E2034 As String = "本予約は既に請求登録が行われているため、取消できません。"
    Public Const E2035 As String = "{0}は単一選択のみ可能です。"
    Public Const E2036 As String = "入力されたユーザIDまたはパスワードが正しくありません。"
    Public Const E2037 As String = "分割差額が存在するため、入力完了できません。"
    Public Const E2038 As String = "編集中のデータが存在する為、帳票を出力できません。"
    Public Const E2039 As String = "入金予定日に土曜日・休日・祝祭日を指定できません。"
    Public Const E2040 As String = "入金情報が入力されています。入金情報を取り消ししてください。"
    Public Const E2041 As String = "EXAS相手先が設定されていないため、請求情報を入力できません。"
    Public Const E2042 As String = "同一付帯設備の情報が設定されています。"
    Public Const E2043 As String = "請求依頼番号が採番されていない請求が選択されています。" & vbCrLf & "EXAS請求データを出力するには請求情報を入力し採番した後に行ってください。"
    Public Const E2044 As String = "請求入力済のデータがある為、金額変更が伴う貸出種別の変更はできません。" & vbCrLf & "変更が必要な場合は、請求入力済のデータを取り消してから再度変更を行ってください。"
    Public Const E2045 As String = "{0}は１件ずつ選択してください。"
    Public Const E2046 As String = "請求情報が入力済のため{0}はできません。" & vbCrLf & "{0}を行う場合は、請求情報を取り消し後に行ってください。"
    Public Const E2047 As String = "対象利用日に該当する料金・倍率のマスタ情報が登録されていません。"
    Public Const E2048 As String = "対象利用日に該当する付帯設備のマスタ情報が登録されていません。" & vbCrLf & "マスタへ登録した後、付帯設備編集処理を行ってください。"
    Public Const E2049 As String = "{0}の合計金額が１２桁を超えています。"
    Public Const E2050 As String = "EXAS相手先が設定されていないため、EXAS用請求依頼ファイルを出力できません。"      '2016.01.26 ADD y.morooka グループ請求対応　EXAS相手先判定
    Public Const E2051 As String = "選択した枠にはすでにキャンセル待ち情報が登録されています。"                      ' 2016.02.19 ADD h.hagiwara 

    '種別:Check
    Public Const C0001 As String = "未定の事務処理がありますが、正式予約登録へ進みますか？"
    Public Const C0002 As String = "入力した内容で{0}の更新を行います。よろしいですか？"
    Public Const C0003 As String = "正式予約に進みます。よろしいですか？"
    Public Const C0004 As String = "入力した内容でキャンセル待ち情報の更新を行います。よろしいですか？"
    Public Const C0005 As String = "仮予約に進みます。よろしいですか？"
    Public Const C0006 As String = "入金取消を行います。よろしいですか？"
    Public Const C0007 As String = "請求取消を行います。よろしいですか？"
    Public Const C0008 As String = "印刷します。よろしいですか？"
    Public Const C0009 As String = "ファイル出力を行います。よろしいですか？"
    Public Const C0010 As String = "承認依頼メールを送信します。よろしいですか？"
    Public Const C0011 As String = "入力した内容で{0}の登録を行います。よろしいですか？"
    Public Const C0012 As String = "{0}と{1}に差異がありますが登録してよろしいでしょうか。"
    Public Const C0013 As String = "EXAS入金データリンクに１行もリンクされていませんが登録してよろしいでしょうか。"
    Public Const C0014 As String = "利用料金に変更が発生します。再設定してください。"
    Public Const C0015 As String = "請求が複数選択されています。よろしいですか？"     '2016.01.15 ADD y.morooka グループ請求対応　チェックボックス複数チェック判定
    Public Const C0016 As String = "{0}に差額があります。このまま入力完了してもよろしいでしょうか。"                       ' 2016.01.27 ADD h.hagiwara 

    '種別:Info
    Public Const I0001 As String = "別途、EXASでの取消操作が必要です。"
    Public Const I0002 As String = "{0}が完了しました。"
    Public Const I0003 As String = "承認記録がありません。（印刷処理は継続されます）"
    Public Const I0004 As String = "アプリケーションを終了します。よろしいですか？"
    Public Const I0005 As String = "別途、EXASでの取消操作が必要です。"

    Public Const EXT_E001 As String = "システムエラーが発生しました。" & vbCrLf & "システム管理者に連絡してください。" & vbCrLf

    ''' <summary>
    ''' プロパティセット【ユーザーID】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrComUserId</returns>
    ''' <remarks><para>作成情報：2015/08/04 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropComStrUserId()
        Get
            Return ppStrComUserId
        End Get
        Set(ByVal value)
            ppStrComUserId = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【ユーザー名】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrComUserName</returns>
    ''' <remarks><para>作成情報：2015/08/04 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropComStrUserName()
        Get
            Return ppStrComUserName
        End Get
        Set(ByVal value)
            ppStrComUserName = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【メールアドレス】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrComMailAddr</returns>
    ''' <remarks><para>作成情報：2015/08/04 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrComMailAddr()
        Get
            Return ppStrComMailAddr
        End Get
        Set(ByVal value)
            ppStrComMailAddr = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【部署コード】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrComBusho</returns>
    ''' <remarks><para>作成情報：2015/08/04 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrComBusho()
        Get
            Return ppStrComBusho
        End Get
        Set(ByVal value)
            ppStrComBusho = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【承認フラグ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrComFlgShonin</returns>
    ''' <remarks><para>作成情報：2015/08/04 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrComFlgShonin()
        Get
            Return ppStrComFlgShonin
        End Get
        Set(ByVal value)
            ppStrComFlgShonin = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【マスタフラグ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrComFlgMst</returns>
    ''' <remarks><para>作成情報：2015/08/04 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrComFlgMst()
        Get
            Return ppStrComFlgMst
        End Get
        Set(ByVal value)
            ppStrComFlgMst = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【シアター定員】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppIntTeiinA</returns>
    ''' <remarks><para>作成情報：2015/08/12 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropIntTeiinA()
        Get
            Return ppIntTeiinA
        End Get
        Set(ByVal value)
            ppIntTeiinA = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【シアター定員】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppIntTeiinB</returns>
    ''' <remarks><para>作成情報：2015/08/12 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropIntTeiinB()
        Get
            Return ppIntTeiinB
        End Get
        Set(ByVal value)
            ppIntTeiinB = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【仮予約事前通知日】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppIntKariJizenTuti</returns>
    ''' <remarks><para>作成情報：2015/08/12 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropIntKariJizenTuti()
        Get
            Return ppIntKariJizenTuti
        End Get
        Set(ByVal value)
            ppIntKariJizenTuti = value
        End Set
    End Property
    ' 2016.01.15 ADD START↓ y.morooka グループ請求対応
    ''' <summary>
    ''' プロパティセット【グループ請求】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppGSeikyusaki</returns>
    ''' <remarks><para>作成情報：2016.01.15 y.morooka
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropGSeikyusaki()
        Get
            Return ppGSeikyusaki
        End Get
        Set(ByVal value)
            ppGSeikyusaki = value
        End Set
    End Property
    ' 2016.01.15 ADD END↑ y.morooka グループ請求対応

    ' 2016.08.12 ADD START↓ m.hayabuchi 請求依頼データ代行処理対応
    ''' <summary>
    ''' プロパティセット【代行フラグ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrDaikoFlg</returns>
    ''' <remarks><para>作成情報：2016.08.12 m.hayabuchi
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrDaikoFlg()
        Get
            Return ppStrDaikoFlg
        End Get
        Set(ByVal value)
            ppStrDaikoFlg = value
        End Set
    End Property
    ' 2016.08.12 ADD END↑ m.hayabuchi 請求依頼データ代行処理対応

    ' 2016.09.20 ADD START↓ m.hayabuchi 請求依頼データ代行処理対応
    ''' <summary>
    ''' プロパティセット【代行担当】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrDaikoTanto</returns>
    ''' <remarks><para>作成情報：2016.09.20 m.hayabuchi
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrDaikoTanto()
        Get
            Return ppStrDaikoTanto
        End Get
        Set(ByVal value)
            ppStrDaikoTanto = value
        End Set
    End Property
    ' 2016.09.20 ADD END↑ m.hayabuchi 請求依頼データ代行処理対応

    ' 2016.09.20 ADD START↓ m.hayabuchi 請求依頼データ代行処理対応
    ''' <summary>
    ''' プロパティセット【代行部署】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrDaikoBusho</returns>
    ''' <remarks><para>作成情報：2016.09.20 m.hayabuchi
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrDaikoBusho()
        Get
            Return ppStrDaikoBusho
        End Get
        Set(ByVal value)
            ppStrDaikoBusho = value
        End Set
    End Property
    ' 2016.09.20 ADD END↑ m.hayabuchi 請求依頼データ代行処理対応

    ''' <summary>
    ''' プロパティセット【フォーム背景色：白】読み取り専用
    ''' </summary>
    ''' <value></value>
    ''' <returns>PropBackColorWhite</returns>
    ''' <remarks><para>作成情報： 2015.11.18 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public ReadOnly Property PropBackColorHONBAN()
        Get
            Return FORM_BACKCOLOR_HONBAN
        End Get
    End Property

    ''' <summary>
    ''' プロパティセット【フォーム背景色：青】読み取り専用
    ''' </summary>
    ''' <value></value>
    ''' <returns>PropBackColorBlue</returns>
    ''' <remarks><para>作成情報： 2015.11.18 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public ReadOnly Property PropBackColorKENSHOU()
        Get
            Return FORM_BACKCOLOR_KENSHOU
        End Get
    End Property

    ''' <summary>
    ''' プロパティセット【環境設定フラグ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppConfigrationFlg</returns>
    ''' <remarks><para>作成情報：2015/08/12 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropConfigrationFlg() As String
        Get
            Return ppConfigrationFlg
        End Get
        Set(ByVal value As String)
            ppConfigrationFlg = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【メインＤＢ】
    ''' </summary>
    ''' <value>メインＤＢ</value>
    ''' <returns>ppMainDB</returns>
    ''' <remarks><para>作成情報：2015/11/18 e.okamura
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropMainDB() As String
        Get
            Return ppMainDB
        End Get
        Set(ByVal value As String)
            ppMainDB = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【メインＤＢダミーファイルパス１】
    ''' </summary>
    ''' <value>メインＤＢダミーファイルパス１</value>
    ''' <returns>ppDummyFilePath1</returns>
    ''' <remarks><para>作成情報：2015/11/18 e.okamura
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDummyFilePath1() As String
        Get
            Return ppDummyFilePath1
        End Get
        Set(ByVal value As String)
            ppDummyFilePath1 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【メインＤＢダミーファイルパス２】
    ''' </summary>
    ''' <value>メインＤＢダミーファイルパス２</value>
    ''' <returns>ppDummyFilePath2</returns>
    ''' <remarks><para>作成情報：2015/11/18 e.okamura
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDummyFilePath2() As String
        Get
            Return ppDummyFilePath2
        End Get
        Set(ByVal value As String)
            ppDummyFilePath2 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【メインＤＢダミーファイル名】
    ''' </summary>
    ''' <value>メインＤＢダミーファイル名</value>
    ''' <returns>ppDummyFileName</returns>
    ''' <remarks><para>作成情報：2015/11/18 e.okamura
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDummyFileName() As String
        Get
            Return ppDummyFileName
        End Get
        Set(ByVal value As String)
            ppDummyFileName = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【ＤＢ接続文字列１】
    ''' </summary>
    ''' <value>ＤＢ接続文字列１</value>
    ''' <returns>ppDBString1</returns>
    ''' <remarks><para>作成情報：2015/11/18 e.okamura
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDBString1() As String
        Get
            Return ppDBString1
        End Get
        Set(ByVal value As String)
            ppDBString1 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【ＤＢ接続文字列２】
    ''' </summary>
    ''' <value>ＤＢ接続文字列２</value>
    ''' <returns>ppDBString2</returns>
    ''' <remarks><para>作成情報：2015/11/18 e.okamura
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDBString2() As String
        Get
            Return ppDBString2
        End Get
        Set(ByVal value As String)
            ppDBString2 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【NetUseユーザーID１】
    ''' </summary>
    ''' <value>NetUseユーザーID１</value>
    ''' <returns>ppNetUseUserID1</returns>
    ''' <remarks><para>作成情報：2015/11/18 e.okamura
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropNetUseUserID1() As String
        Get
            Return ppNetUseUserID1
        End Get
        Set(ByVal value As String)
            ppNetUseUserID1 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【NetUseパスワード１】
    ''' </summary>
    ''' <value>NetUseパスワード１</value>
    ''' <returns>ppNetUsePassword1</returns>
    ''' <remarks><para>作成情報：2015/11/18 e.okamura
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropNetUsePassword1() As String
        Get
            Return ppNetUsePassword1
        End Get
        Set(ByVal value As String)
            ppNetUsePassword1 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【NetUseユーザーID２】
    ''' </summary>
    ''' <value>NetUseユーザーID２</value>
    ''' <returns>ppNetUseUserID2</returns>
    ''' <remarks><para>作成情報：2015/11/18 e.okamura
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropNetUseUserID2() As String
        Get
            Return ppNetUseUserID2
        End Get
        Set(ByVal value As String)
            ppNetUseUserID2 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【NetUseパスワード２】
    ''' </summary>
    ''' <value>NetUseパスワード２</value>
    ''' <returns>ppNetUsePassword2</returns>
    ''' <remarks><para>作成情報：2015/11/18 e.okamura
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropNetUsePassword2() As String
        Get
            Return ppNetUsePassword2
        End Get
        Set(ByVal value As String)
            ppNetUsePassword2 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【ＤＢチェック有無】
    ''' </summary>
    ''' <value>ＤＢチェック有無</value>
    ''' <returns>ppDBCheck</returns>
    ''' <remarks><para>作成情報：2015/11/18 e.okamura
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDBCheck() As String
        Get
            Return ppDBCheck
        End Get
        Set(ByVal value As String)
            ppDBCheck = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【Debug用マスタＤＢ】
    ''' </summary>
    ''' <value>Debug用マスタＤＢ</value>
    ''' <returns>ppDebugDB</returns>
    ''' <remarks><para>作成情報：2015/11/18 e.okamura
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDebugDB() As String
        Get
            Return ppDebugDB
        End Get
        Set(ByVal value As String)
            ppDebugDB = value
        End Set
    End Property

End Module