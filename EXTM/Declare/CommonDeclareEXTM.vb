Public Module CommonDeclareEXTM
    Public Const E0000 As String = "システムエラーが発生しました。システム管理者に連絡してください。"
    Public Const E0001 As String = "{0}を入力してください。"
    Public Const E0002 As String = "{0}を選択してください。"
    Public Const E0003 As String = "{0}は半角数字で入力してください。"
    Public Const E0005 As String = "{0}は半角英数字で入力してください。"
    Public Const E0009 As String = "{0}は{1}桁以上{2}桁以内で入力してください。"
    Public Const E0010 As String = "{0}は{1}桁で入力してください。"
    Public Const E0018 As String = "{0}は{1}から{2}の範囲内で入力してください。"
    Public Const E0020 As String = "{0}は{1}より小さい値を入力してください。"

    Public Const E2032 As String = "入力された期間は設定済みのため登録できません。"

    Public Const C0011 As String = "入力した内容で{0}の登録を行います。よろしいですか？"

    Public Const I0002 As String = "{0}が完了しました。"

    '共通エラー表示
    Public Const EXTM0102_E0000 As String = "システムエラーが発生しました。システム管理者に連絡してください。"

    '期間、分類、付帯設備表未入力チェック
    Public Const EXTM0102_E0001 As String = "{0}が入力されていません。"

    ' コンボ入力チェック
    Public Const EXTM0102_E0002 As String = "{0}が選択されていません。"

    '数字のみ入力
    Public Const EXTM0102_E0003 As String = "{0}は半角数字で入力してください。"

    '桁数チェック
    Public Const EXTM0102_E0010 As String = "{0}は{1}桁で入力してください。"

    '期間月チェック
    Public Const EXTM0102_E0018 As String = "{0}は{1}から{2}の範囲内で入力してください。"

    '期間逆転チェック
    Public Const EXTM0102_E2011 As String = "{0}は{1}以降の日付を入力してください。"

    '期間登録済みチェック
    Public Const EXTM0102_E2032 As String = "入力された期間は設定済みのため登録できません。"

    '登録確認メッセージ
    Public Const EXTM0102_C0011 As String = "入力した内容で{0}の登録を行います。よろしいですか？"

    '登録完了メッセージ
    Public Const EXTM0102_C0012 As String = "{0}の登録が完了しました。"

    'EXAS相手先一覧(EXTM0203)メッセージ
    '入力チェック（０件）
    Public Const EXTM0203_E0002 As String = "{0}を選択してください。"

    '入力チェック（２件以上）
    Public Const EXTM0203_E2035 As String = "{0}は単一選択のみ可能です。"
    Public Const EXTM0203_E2036 As String = "検索条件を入力してください。"

    '利用者一覧一覧画面（EXTM0201）
    'エラーメッセージ
    Public Const M0201_E0000 As String = "システムエラーが発生しました。" & vbCrLf & "システム管理者に連絡してください。" & vbCrLf
    Public Const M0201_E0002 As String = "{0}を選択してください。"
    Public Const M0201_E2035 As String = "{0}は単一選択のみ可能です。"
    'メッセージボックスタイトル
    Public Const TITLE_ERROR As String = "エラー"
    ' スプレッド位置
    Public Const M0201_COL_RIYO_SELECT As Integer = 0
    Public Const M0201_COL_RIYO_CD As Integer = 1
    Public Const M0201_COL_RIYO_LVL As Integer = 2
    Public Const M0201_COL_RIYO_NM As Integer = 3
    Public Const M0201_COL_RIYO_KNM As Integer = 4
    Public Const M0201_COL_RIYO_LAST_DAY As Integer = 5
    Public Const M0201_COL_RIYO_AITE_CD As Integer = 6
    Public Const M0201_COL_RIYO_AITE_NM As Integer = 7
    Public Const M0201_COL_RIYO_TEL1 As Integer = 8
    Public Const M0201_COL_RIYO_TEL2 As Integer = 9
    Public Const M0201_COL_RIYO_FAX As Integer = 10
    Public Const M0201_COL_RIYO_YUBIN As Integer = 11
    Public Const M0201_COL_RIYO_ADDRESS As Integer = 12
    Public Const M0201_COL_RIYO_BUTTON As Integer = 13
    Public Const M0201_COL_RIYO_SLVL As Integer = 14
    Public Const M0201_COL_RIYO_SDAIHYO_NM As Integer = 15
    Public Const M0201_COL_RIYO_STEL11 As Integer = 16
    Public Const M0201_COL_RIYO_STEL12 As Integer = 17
    Public Const M0201_COL_RIYO_STEL13 As Integer = 18
    Public Const M0201_COL_RIYO_SNAISEN As Integer = 19
    Public Const M0201_COL_RIYO_SFAX11 As Integer = 20
    Public Const M0201_COL_RIYO_SFAX12 As Integer = 21
    Public Const M0201_COL_RIYO_SFAX13 As Integer = 22
    Public Const M0201_COL_RIYO_SYUBIN1 As Integer = 23
    Public Const M0201_COL_RIYO_SYUBIN2 As Integer = 24
    Public Const M0201_COL_RIYO_ADD1 As Integer = 25
    Public Const M0201_COL_RIYO_ADD2 As Integer = 26
    Public Const M0201_COL_RIYO_ADD3 As Integer = 27
    Public Const M0201_COL_RIYO_ADD4 As Integer = 28


    '利用者情報登録/詳細画面
    'エラーメッセージ
    Public Const M0202_E0001 As String = "{0}を入力してください。"
    Public Const M0202_E0002 As String = "{0}を選択してください。"
    Public Const M0202_E0003 As String = "{0}は半角数字で入力してください。"
    Public Const M0202_E0005 As String = "{0}は半角英数字で入力してください。"
    Public Const M0202_E0006 As String = "{0}は全角文字で入力してください。"
    Public Const M0202_E0007 As String = "{0}は全角カナで入力してください。"
    Public Const M0202_E0009 As String = "{0}は{1}桁以上{2}桁以内で入力してください。"
    Public Const M0202_C0011 As String = "入力した内容で{0}の登録を行います。よろしいですか？"
    Public Const M0202_I0002 As String = "{0}が完了しました。"

    '都道府県コンボボックス用配列
    Public strCmbToDo(,) As String = {{"-1", "選択してください"},
                                      {"0", "北海道"}, {"1", "青森県"},
                                      {"2", "岩手県"}, {"3", "宮城県"},
                                      {"4", "秋田県"}, {"5", "山形県"},
                                      {"6", "福島県"}, {"7", "茨城県"},
                                      {"8", "栃木県"}, {"9", "群馬県"},
                                      {"10", "埼玉県"}, {"11", "千葉県"},
                                      {"12", "東京都"}, {"13", "神奈川県"},
                                      {"14", "新潟県"}, {"15", "富山県"},
                                      {"16", "石川県"}, {"17", "福井県"},
                                      {"18", "山梨県"}, {"19", "長野県"},
                                      {"20", "岐阜県"}, {"21", "静岡県"},
                                      {"22", "愛知県"}, {"23", "三重県"},
                                      {"24", "滋賀県"}, {"25", "京都府"},
                                      {"26", "大阪府"}, {"27", "兵庫県"},
                                      {"28", "奈良県"}, {"29", "和歌山県"},
                                      {"30", "鳥取県"}, {"31", "島根県"},
                                      {"32", "岡山県"}, {"33", "広島県"},
                                      {"34", "山口県"}, {"35", "徳島県"},
                                      {"36", "香川県"}, {"37", "愛媛県"},
                                      {"38", "高知県"}, {"39", "福岡県"},
                                      {"40", "佐賀県"}, {"41", "長崎県"},
                                      {"42", "熊本県"}, {"43", "大分県"},
                                      {"44", "宮崎県"}, {"45", "鹿児島県"},
                                      {"46", "沖縄県"}}



    'EXシアターユーザーマスタメンテ一覧画面（EXTM0101）
    Public Const M0101_E0000 As String = "システムエラーが発生しました。システム管理者に連絡してください。"
    Public Const M0101_E0001 As String = "{0}を入力してください。"
    Public Const M0101_E0005 As String = "{0}は半角英数字で入力してください。"
    Public Const M0101_E0008 As String = "入力された{0}は不正なアドレスのようです。再度入力し直してください。"
    Public Const M0101_E0009 As String = "{0}は{1}桁以上{2}桁以内で入力してください。"
    Public Const M0101_E0012 As String = "{0}に入力された値が正しくありません。"
    Public Const M0101_E1002 As String = "入力された{0}は既に使用されています。"
    Public Const M0101_E0014 As String = "{0}を一つ以上選択してください。"
    Public Const M0101_C0011 As String = "入力した内容で{0}の登録を行います。よろしいですか？"
    Public Const M0101_I0002 As String = "{0}が完了しました。"

    'EXシアター消費税マスタメンテ一覧画面（EXTM0103）
    Public Const M0103_E0000 As String = "システムエラーが発生しました。システム管理者に連絡してください。"
    Public Const M0103_E0001 As String = "{0}を入力してください。"
    Public Const M0103_E0003 As String = "{0}は半角数字で入力してください。"
    Public Const M0103_E0004 As String = "{0}は半角数字または記号で入力してください。"
    Public Const M0103_E0009 As String = "{0}は{1}桁以上{2}桁以内で入力してください。"
    Public Const M0103_E2005 As String = "{0}の日付は正しくありません。"
    Public Const M0103_E2011 As String = "{0}は{1}以降の日付を入力してください。"
    Public Const M0103_E2032 As String = "入力された期間は設定済みのため登録できません。"
    Public Const M0103_E2039 As String = "{0}を変更してください。"
    Public Const M0103_E2040 As String = "最終行の{0}は空白のみ可能です。"
    Public Const M0103_C0011 As String = "入力した内容で{0}の登録を行います。よろしいですか？"
    Public Const M0103_I0002 As String = "{0}が完了しました。"

End Module
