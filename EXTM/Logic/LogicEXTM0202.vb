Imports Common
Imports CommonEXT
Imports Npgsql
Imports FarPoint.Win.Spread

''' <summary>利用者情報登録/詳細画面Logicクラス
''' </summary>
''' <remarks>利用者情報登録/詳細画面Logicクラスを定義する
''' <para>作成情報：2015/08/25 ozawa
''' <p>改訂情報：</p>
''' </para> </remarks>
Public Class LogicEXTM0202
    'インスタンス生成 
    Private SqlEXTM0202 As New SqlEXTM0202          'SQLクラス
    Private CommonLogic As New CommonLogic          'CommonLogicクラス
    Private CommonLogicEXT As New CommonLogicEXT    'CommonLogicEXTクラス
    Public DataEXTM0202 As New DataEXTM0202         'Dataクラス
    Public CommonValidation As CommonValidation     'CommonValidation

    ''' <summary>初期表示時、利用者情報を取得する
    ''' </summary>
    ''' <param name="dataEXTM0202"></param>
    ''' <returns></returns>
    ''' <remarks>条件に応じて利用者情報を取得する
    ''' <para>作成情報：2015/08/25 ozawa]
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function GetRiyoshaData(ByRef dataEXTM0202 As DataEXTM0202)
        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)       'コネクション
        Dim Adapter As New NpgsqlDataAdapter           'アダプター
        Dim Table As New DataTable()                   'テーブル

        Try
            'コネクションを開く
            Cn.Open()
            'エラーメッセージ初期化
            puErrMsg = System.String.Empty

            'SELECT用SQLコマンドを作成
            If SqlEXTM0202.GetRiyoshaData(Adapter, Cn, dataEXTM0202) = False Then
                Return False
            End If

            'SQLログ
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "利用者情報検索", Nothing, Adapter.SelectCommand)

            'データを取得
            Adapter.Fill(Table)

            '取得データをDataクラスへ保存
            'テキストボックスに入れる
            dataEXTM0202.PropDtRiyoshaMasta = Table

            '取得したデータをフォームオブジェクトに格納

            dataEXTM0202.PropTxtRiyo_nm.Text = dataEXTM0202.PropDtRiyoshaMasta.Rows(0).Item(1)           '利用者名
            dataEXTM0202.PropTxtRiyo_kana.Text = dataEXTM0202.PropDtRiyoshaMasta.Rows(0).Item(2)         '利用者カナ
            dataEXTM0202.PropTxtDaihyo_nm.Text = dataEXTM0202.PropDtRiyoshaMasta.Rows(0).Item(3)
            dataEXTM0202.PropTxtTel1.Text = dataEXTM0202.PropDtRiyoshaMasta.Rows(0).Item(4)              '電話番号1
            dataEXTM0202.PropTxtTel2.Text = dataEXTM0202.PropDtRiyoshaMasta.Rows(0).Item(5)              '電話番号2
            dataEXTM0202.PropTxtTel3.Text = dataEXTM0202.PropDtRiyoshaMasta.Rows(0).Item(6)              '電話番号3
            dataEXTM0202.PropTxtNaisen.Text = dataEXTM0202.PropDtRiyoshaMasta.Rows(0).Item(7)            '内線番号
            dataEXTM0202.PropTxtFax1.Text = dataEXTM0202.PropDtRiyoshaMasta.Rows(0).Item(8)              'Fax番号1
            dataEXTM0202.PropTxtFax2.Text = dataEXTM0202.PropDtRiyoshaMasta.Rows(0).Item(9)              'Fax番号2
            dataEXTM0202.PropTxtFax3.Text = dataEXTM0202.PropDtRiyoshaMasta.Rows(0).Item(10)              'Fax番号3
            dataEXTM0202.PropTxtYubin1.Text = dataEXTM0202.PropDtRiyoshaMasta.Rows(0).Item(11)            '郵便番号1
            dataEXTM0202.PropTxtYubin2.Text = dataEXTM0202.PropDtRiyoshaMasta.Rows(0).Item(12)            '郵便番号2

            'コンボボックス
            dataEXTM0202.PropCmbTodo.Text = dataEXTM0202.PropDtRiyoshaMasta.Rows(0).Item(13)             '都道府県
            dataEXTM0202.PropTxtShiku.Text = dataEXTM0202.PropDtRiyoshaMasta.Rows(0).Item(14)             '市区町村
            dataEXTM0202.PropTxtBanchi.Text = dataEXTM0202.PropDtRiyoshaMasta.Rows(0).Item(15)            '番地
            dataEXTM0202.PropTxtBuild.Text = dataEXTM0202.PropDtRiyoshaMasta.Rows(0).Item(16)             'ビル名
            
            'ラジオボタン表示　利用レベル：1の場合
            If dataEXTM0202.PropDtRiyoshaMasta.Rows(0).Item(17) = "1" Then
                '「通常」にセット
                dataEXTM0202.PropRdoTujyo.Checked = True
                '利用レベル：2の場合
            ElseIf dataEXTM0202.PropDtRiyoshaMasta.Rows(0).Item(17) = "2" Then
                '「要注意」にセット
                dataEXTM0202.PropRdoChui.Checked = True
                '利用レベル：3の場合
            ElseIf dataEXTM0202.PropDtRiyoshaMasta.Rows(0).Item(17) = "3" Then
                '「利用不可」にセット
                dataEXTM0202.PropRdoHuka.Checked = True
            End If

            dataEXTM0202.PropLblAite_cd.Text = dataEXTM0202.PropDtRiyoshaMasta.Rows(0).Item(18)           '相手先コード
            dataEXTM0202.PropLblAite_nm.Text = dataEXTM0202.PropDtRiyoshaMasta.Rows(0).Item(19)           '相手先名
            dataEXTM0202.PropTxtCom.Text = dataEXTM0202.PropDtRiyoshaMasta.Rows(0).Item(20)               'コメント


            '終了ログ
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            Return True

        Catch ex As Exception
            '例外処理
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            puErrMsg = M0201_E0000 & ex.Message
            Return False

        Finally
            Cn.Close()
            Cn.Dispose()
            Adapter.Dispose()
            Table.Dispose()

        End Try

    End Function

    ''' <summary>初期表示時、利用コメントを取得する
    ''' </summary>
    ''' <param name="dataEXTM0202"></param>
    ''' <returns></returns>
    ''' <remarks>条件に応じて利用コメントを取得する
    ''' <para>作成情報：2015/08/25 ozawa]
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Function GetComment(ByRef dataEXTM0202 As DataEXTM0202)
        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)
        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)       'コネクション
        Dim Adapter As New NpgsqlDataAdapter           'アダプター
        Dim Table As New DataTable()                   'テーブル

        Try
            'コネクションを開く
            Cn.Open()
            'エラーメッセージ初期化
            puErrMsg = System.String.Empty

            'SELECT用SQLコマンドを作成
            If SqlEXTM0202.GetComment(Adapter, Cn, dataEXTM0202) = False Then
                Return False
            End If

            'SQLログ
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "利用コメント検索", Nothing, Adapter.SelectCommand)

            'データを取得
            Adapter.Fill(Table)

            '取得データをDataクラスへ保存
            dataEXTM0202.PropVwList.DataSource = Table

            '0件の場合
            '閾値を超えている場合


            '終了ログ
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)
            Return True

        Catch ex As Exception
            '例外処理
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, ex, Nothing)
            puErrMsg = M0201_E0000 & ex.Message
            Return False

        Finally
            Cn.Close()
            Cn.Dispose()
            Adapter.Dispose()
            Table.Dispose()

        End Try

    End Function

    ''' <summary>
    ''' 入力チェックメイン処理
    ''' </summary>
    ''' <param name="dataEXTM0202"></param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>登録項目の入力チェックを行う
    ''' <para>作成情報：2015/08/25 ozawa
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function CheckInputValueMain(ByRef dataEXTM0202 As DataEXTM0202) As Boolean
        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        'コントロール入力チェック
        If CheckInputValue(dataEXTM0202) = False Then
            Return False
        End If

        '終了ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

        '正常処理終了
        Return True

    End Function

    ''' <summary>
    ''' 入力チェック処理
    ''' </summary>
    ''' <param name="dataEXTM0202"></param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>登録項目の入力チェックを行う
    ''' <para>作成情報：2015/08/25 ozawa
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Private Function CheckInputValue(ByRef dataEXTM0202 As DataEXTM0202) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        Try
            Dim CommonValidation As New CommonValidation

            With dataEXTM0202

                '利用者名
                With .PropTxtRiyo_nm
                    '未入力の場合、エラー
                    If .Text.Trim() = "" Then
                        'エラーメッセージ設定
                        puErrMsg = String.Format(M0202_E0001, "利用者名")
                        MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                        Return False
                    End If
                    '最大値
                    'If .Text.Length > 20 Then            ' 2015.10.06 UPDATE h.hagiwara 桁数変更(20⇒30)
                    If .Text.Length > 30 Then             ' 2015.10.06 UPDATE h.hagiwara 桁数変更(20⇒30)
                        'puErrMsg = String.Format(M0202_E0009, "利用者名", "1", "20")            ' 2015.10.06 UPDATE h.hagiwara 桁数変更(20⇒30)
                        puErrMsg = String.Format(M0202_E0009, "利用者名", "1", "20")             ' 2015.10.06 UPDATE h.hagiwara 桁数変更(20⇒30)
                        MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                        Return False
                    End If

                    '全角チェック
                    If CommonValidation.IsFullChar(.Text) = False Then
                        puErrMsg = String.Format(M0202_E0006, "利用者名")
                        MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                        Return False
                    End If
                End With

                '利用者カナ
                With .PropTxtRiyo_kana
                    '未入力の場合、エラー
                    If .Text.Trim() = "" Then
                        'エラーメッセージ設定
                        puErrMsg = String.Format(M0202_E0001, "利用者名（カナ）")
                        MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                        Return False
                    End If
                    '最大値
                    If .Text.Length > 30 Then
                        puErrMsg = String.Format(M0202_E0009, "利用者名（カナ）", "1", "30")
                        MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                        Return False
                    End If
                    '全角カナチェック
                    'If IsFullKana(dataEXTM0202.PropTxtRiyo_kana.Text) = False Then                         ' 2015.11.20 UPD h.hagiwara
                    If CommonLogicEXT.IsFullKana(dataEXTM0202.PropTxtRiyo_kana.Text) = False Then           ' 2015.11.20 UPD h.hagiwara
                        puErrMsg = String.Format(M0202_E0007, "利用者名（カナ）")
                        MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                        Return False
                    End If
                End With

                '代表者名
                With .PropTxtDaihyo_nm
                    '未入力の場合、エラー
                    If .Text.Trim() <> "" Then
                    '最大値
                    If .Text.Length > 20 Then
                        puErrMsg = String.Format(M0202_E0009, "代表者名", "1", "20")
                        MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                        Return False
                    End If
                    '全角チェック
                        If CommonValidation.IsFullChar(dataEXTM0202.PropTxtDaihyo_nm.Text) = False Then
                            puErrMsg = String.Format(M0202_E0006, "代表者名")
                            MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                            Return False
                        End If
                    End If
                End With

                '電話番号1
                With .PropTxtTel1
                    '未入力の場合、エラー
                    If .Text.Trim() = "" Then
                        'エラーメッセージ設定
                        puErrMsg = String.Format(M0202_E0001, "電話番号1")
                        MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                        Return False
                    End If
                    '最大値
                    If .Text.Length > 6 Then
                        puErrMsg = String.Format(M0202_E0009, "電話番号1", "1", "6")
                        MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                        Return False
                    End If
                    '半角数字チェック
                    If CommonValidation.IsHalfNmb(dataEXTM0202.PropTxtTel1.Text) = False Then
                        puErrMsg = String.Format(M0202_E0003, "電話番号1")
                        MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                        Return False
                    End If
                End With

                '電話番号2
                With .PropTxtTel2
                    '未入力の場合、エラー
                    If .Text.Trim() = "" Then
                        'エラーメッセージ設定
                        puErrMsg = String.Format(M0202_E0001, "電話番号2")
                        MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                        Return False
                    End If
                    '最大値
                    If .Text.Length > 6 Then
                        puErrMsg = String.Format(M0202_E0009, "電話番号2", "1", "6")
                        MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                        Return False
                    End If
                    '半角数字チェック
                    If CommonValidation.IsHalfNmb(dataEXTM0202.PropTxtTel2.Text) = False Then
                        puErrMsg = String.Format(M0202_E0003, "電話番号2")
                        MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                        Return False
                    End If
                End With

                '電話番号3
                With .PropTxtTel3
                    '未入力の場合、エラー
                    If .Text.Trim() = "" Then
                        'エラーメッセージ設定
                        puErrMsg = String.Format(M0202_E0001, "電話番号3")
                        MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                        Return False
                    End If
                    '最大値
                    If .Text.Length > 6 Then
                        puErrMsg = String.Format(M0202_E0009, "電話番号3", "1", "6")
                        MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                        Return False
                    End If
                    '半角数字チェック
                    If CommonValidation.IsHalfNmb(dataEXTM0202.PropTxtTel3.Text) = False Then
                        puErrMsg = String.Format(M0202_E0003, "電話番号3")
                        MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                        Return False
                    End If
                End With

                '内線番号
                With .PropTxtNaisen
                    '必須項目ではないため、入力されている場合のみチェックする
                    If .Text.Trim() <> "" Then
                        '最大値
                        If .Text.Length > 6 Then
                            puErrMsg = String.Format(M0202_E0009, "内線", "1", "6")
                            MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                            Return False
                        End If
                        '半角数字チェック
                        If CommonValidation.IsHalfNmb(dataEXTM0202.PropTxtNaisen.Text) = False Then
                            puErrMsg = String.Format(M0202_E0003, "内線")
                            MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                            Return False
                        End If
                    End If

                End With

                'Fax番号1
                With .PropTxtFax1
                    '必須項目ではないため、入力されている場合のみ入力チェックする
                    If .Text.Trim() <> "" Then
                        '最大値
                        If .Text.Length > 6 Then
                            puErrMsg = String.Format(M0202_E0009, "Fax1", "1", "6")
                            MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                            Return False
                        End If
                        '半角数字チェック
                        If CommonValidation.IsHalfNmb(dataEXTM0202.PropTxtFax1.Text) = False Then
                            puErrMsg = String.Format(M0202_E0003, "Fax1")
                            MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                            Return False
                        End If
                        '項目間チェック
                        If dataEXTM0202.PropTxtFax2.Text.Trim() = "" Then
                            puErrMsg = String.Format(M0202_E0001, "Fax2")
                            MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                            Return False
                        ElseIf dataEXTM0202.PropTxtFax3.Text.Trim() = "" Then
                            puErrMsg = String.Format(M0202_E0001, "Fax3")
                            MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                            Return False
                        End If
                    End If
                End With
                'Fax番号2
                With .PropTxtFax2
                    '必須項目ではないため、入力されている場合のみ入力チェックする
                    If .Text.Trim() <> "" Then
                        '最大値
                        If .Text.Length > 6 Then
                            puErrMsg = String.Format(M0202_E0009, "Fax2", "1", "6")
                            MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                            Return False
                        End If
                        '半角数字チェック
                        If CommonValidation.IsHalfNmb(dataEXTM0202.PropTxtFax2.Text) = False Then
                            puErrMsg = String.Format(M0202_E0003, "Fax2")
                            MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                            Return False
                        End If
                        '項目間チェック
                        If dataEXTM0202.PropTxtFax1.Text.Trim() = "" Then
                            puErrMsg = String.Format(M0202_E0001, "Fax1")
                            MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                            Return False
                        ElseIf dataEXTM0202.PropTxtFax3.Text.Trim() = "" Then
                            puErrMsg = String.Format(M0202_E0001, "Fax3")
                            MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                            Return False
                        End If
                    End If
                End With

                'Fax番号3
                With .PropTxtFax3
                    '必須項目ではないため、入力されている場合のみ入力チェックする
                    If .Text.Trim() <> "" Then
                        '最大値
                        If .Text.Length > 6 Then
                            puErrMsg = String.Format(M0202_E0009, "Fax3", "1", "6")
                            MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                            Return False
                        End If
                        '半角数字チェック
                        If CommonValidation.IsHalfNmb(dataEXTM0202.PropTxtFax3.Text) = False Then
                            puErrMsg = String.Format(M0202_E0003, "Fax3")
                            MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                            Return False
                        End If
                        '項目間チェック
                        If dataEXTM0202.PropTxtFax1.Text.Trim() = "" Then
                            puErrMsg = String.Format(M0202_E0001, "Fax1")
                            MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                            Return False
                        ElseIf dataEXTM0202.PropTxtFax2.Text.Trim() = "" Then
                            puErrMsg = String.Format(M0202_E0001, "Fax2")
                            MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                            Return False
                        End If
                    End If
                End With

                '郵便番号1
                With .PropTxtYubin1
                    '未入力の場合、エラー
                    If .Text.Trim() = "" Then
                        'エラーメッセージ設定
                        puErrMsg = String.Format(M0202_E0001, "郵便番号1")
                        MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                        Return False
                    End If
                    ' 2015.12.07 UPD START↓ h.hagiwara 最小値も判定
                    '最大値
                    'If .Text.Length > 3 Then
                    '    puErrMsg = String.Format(M0202_E0009, "郵便番号1", "1", "3")
                    '    MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                    '    Return False
                    'End If
                    If .Text.Length > 3 Or .Text.Length < 3 Then
                        puErrMsg = String.Format(E0010, "郵便番号1", "3")
                        MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                        Return False
                    End If
                    ' 2015.12.07 UPD END↑ h.hagiwara 最小値も判定
                    '半角数字チェック
                    If CommonValidation.IsHalfNmb(dataEXTM0202.PropTxtYubin1.Text) = False Then
                        puErrMsg = String.Format(M0202_E0003, "郵便番号1")
                        MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                        Return False
                    End If
                End With

                '郵便番号2
                With .PropTxtYubin2
                    '未入力の場合、エラー
                    If .Text.Trim() = "" Then
                        'エラーメッセージ設定
                        puErrMsg = String.Format(M0202_E0001, "郵便番号2")
                        MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                        Return False
                    End If
                    ' 2015.12.07 UPD START↓ h.hagiwara 最小値も判定
                    '最大値
                    'If .Text.Length > 4 Then
                    '    puErrMsg = String.Format(M0202_E0009, "郵便番号2", "1", "4")
                    '    MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                    '    Return False
                    'End If
                    If .Text.Length > 4 Or .Text.Length < 4 Then
                        puErrMsg = String.Format(E0010, "郵便番号2", "4")
                        MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                        Return False
                    End If
                    ' 2015.12.07 UPD END↑ h.hagiwara 最小値も判定
                    '半角数字チェック
                    If CommonValidation.IsHalfNmb(dataEXTM0202.PropTxtYubin2.Text) = False Then
                        puErrMsg = String.Format(M0202_E0003, "郵便番号2")
                        MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                        Return False
                    End If
                End With

                '都道府県
                With .PropCmbTodo
                    '未入力の場合、エラー
                    If .Text.Trim() = "" Then
                        'エラーメッセージ設定
                        puErrMsg = String.Format(M0202_E0002, "都道府県")
                        MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                        Return False
                    End If
                    '未選択時、エラー
                    If .SelectedIndex = 0 Then
                        'エラーメッセージ設定
                        puErrMsg = String.Format(M0202_E0002, "都道府県")
                        MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                        Return False
                    End If
                End With
                '市区町村
                With .PropTxtShiku
                    '未入力の場合、エラー
                    If .Text.Trim() = "" Then
                        'エラーメッセージ設定
                        puErrMsg = String.Format(M0202_E0001, "市区町村")
                        MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                        Return False
                    End If
                    '最大値
                    If .Text.Length > 20 Then
                        puErrMsg = String.Format(M0202_E0009, "市区町村", "1", "20")
                        MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                        Return False
                    End If
                End With
                '番地
                With .PropTxtBanchi
                    '未入力の場合、エラー
                    If .Text.Trim() = "" Then
                        'エラーメッセージ設定
                        puErrMsg = String.Format(M0202_E0001, "番地")
                        MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                        Return False
                    End If
                    '最大値
                    If .Text.Length > 20 Then
                        puErrMsg = String.Format(M0202_E0009, "番地", "1", "20")
                        MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                        Return False
                    End If
                End With
                'ビル名
                With .PropTxtBuild
                    '必須項目ではないため、入力されている場合のみチェックする
                    If .Text.Trim() <> "" Then
                        '最大値
                        If .Text.Length > 20 Then
                            puErrMsg = String.Format(M0202_E0009, "ビル名", "1", "20")
                            MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                            Return False
                        End If
                    End If
                End With

                '相手先コード
                With .PropLblAite_cd
                    '未入力の場合、エラー
                    If .Text.Trim() = "" Then
                        'エラーメッセージ設定
                        puErrMsg = String.Format(M0202_E0001, "EXAS相手先コード")
                        MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                        Return False
                    End If
                    '最大値
                    If .Text.Length > 6 Then
                        puErrMsg = String.Format(M0202_E0009, "EXAS相手先コード", "1", "6")
                        MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                        Return False
                    End If
                    '半角英数字チェック
                    If CommonValidation.IsHalfChar(dataEXTM0202.PropLblAite_cd.Text) = False Then
                        puErrMsg = String.Format(M0202_E0005, "EXAS相手先コード")
                        MsgBox(puErrMsg, MsgBoxStyle.Critical, TITLE_ERROR)
                        Return False
                    End If
                End With
            End With

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = M0201_E0000 & ex.Message
            Return False
        End Try
    End Function

    ''' <summary>データ更新処理
    ''' </summary>
    ''' <param name="dataEXTM0202"></param>
    ''' <returns>>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>入力内容をDBに反映（UPDATE）する
    ''' <para>作成情報：2015/08/25 ozawa
    ''' <p>改訂情報：</p>
    ''' </para>
    ''' </remarks>
    Public Function UpdatetNewData(ByVal dataEXTM0202 As DataEXTM0202) As Boolean
        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)  'サーバーとクライアントをつなげる
        Dim Adapter As New NpgsqlDataAdapter()    'アダプタ
        Dim Tsx As NpgsqlTransaction = Nothing    'トランザクション

        Try
            'コネクションを開く
            Cn.Open()

            'トランザクションレベルを設定し、トランザクションを開始する
            Tsx = Cn.BeginTransaction(IsolationLevel.Serializable)

            'サーバー日付取得処理
            If SelectSysdate(Adapter, Cn, dataEXTM0202) = False Then
                Return False
            End If
            '利用者情報更新処理
            If UpdateNewRiyoshaData(Tsx, Cn, dataEXTM0202) = False Then
                Return False
            End If

            'コミット
            Tsx.Commit()

            'コネクションを閉じる
            'Cn.Close()

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'ロールバック
            'If Tsx IsNot Nothing Then
            '    Tsx.Rollback()
            'End If
            'コネクションが閉じられていない場合は閉じる
            If Cn IsNot Nothing Then
                Cn.Close()
            End If
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = ex.Message
            Return False
        Finally

            If Tsx IsNot Nothing Then
                Tsx.Dispose()
            End If
            Cn.Dispose()
            Cn.Close()
        End Try
    End Function

    ''' <summary>データ新規登録処理
    ''' </summary>
    ''' <param name="dataEXTM0202"></param>
    ''' <returns>Boolean True:正常終了 False:異常終了</returns>
    ''' <remarks>入力内容をDBに新規登録（INSERT）する
    ''' <para>作成情報：2015/08/25 ozawa
    ''' <p>改訂情報 : </p>
    ''' </para></remarks>
    Public Function InsertNewData(ByVal dataEXTM0202 As DataEXTM0202) As Boolean

        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cn As New NpgsqlConnection(DbString)  'サーバーとクライアントをつなげる
        Dim Tsx As NpgsqlTransaction = Nothing    'トランザクション

        Try
            'コネクションを開く
            Cn.Open()

            'トランザクションレベルを設定し、トランザクションを開始する
            Tsx = Cn.BeginTransaction(IsolationLevel.Serializable)

            '新規利用者番号、システム日付取得
            If GetNewRiyoshaCd(Cn, dataEXTM0202) = False Then
                Return False
            End If

            '利用者情報新規登録
            If InsertNewRiyoshaData(Tsx, Cn, dataEXTM0202) = False Then
                'ロールバック
                Tsx.Rollback()
                Return False
            End If

            'コミット
            Tsx.Commit()

            'コネクションを閉じる
            Cn.Close()

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'ロールバック
            If Tsx IsNot Nothing Then
                Tsx.Rollback()
            End If
            'コネクションが閉じられていない場合は閉じる
            If Cn IsNot Nothing Then
                Cn.Close()
            End If
            'メッセージ変数にエラーメッセージを格納
            'puErrMsg = HBK_E001 & ex.Message
            Return False
        Finally
            If Tsx IsNot Nothing Then
                Tsx.Dispose()
            End If
            Cn.Dispose()
        End Try

    End Function

    ''' <summary>サーバー日付取得処理
    ''' </summary>
    ''' <param name="Adapter"></param>
    ''' <param name="Cn"></param>
    ''' <param name="dataEXTM0202"></param>
    ''' <returns></returns>
    ''' <remarks>サーバー日付を取得する
    ''' <para>作成情報：20015/08/25 ozawa
    ''' <p>改訂情報：</p>
    ''' </para>
    ''' </remarks>
    Public Function SelectSysdate(ByRef Adapter As NpgsqlDataAdapter, _
                                  ByVal Cn As NpgsqlConnection, _
                                  ByVal dataEXTM0202 As DataEXTM0202) As Boolean
        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim dtResult As New DataTable

        Try

            'SQLを作成
            If SqlEXTM0202.SelectSysdate(Adapter, Cn, dataEXTM0202) = False Then
                Return False
            End If

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "新規利用者番号、システム日付取得", Nothing, Adapter.SelectCommand)

            'SQL実行
            Adapter.Fill(dtResult)

            'データが取得できた場合、データクラスに取得データをセット
            If dtResult.Rows.Count > 0 Then
                dataEXTM0202.PropDtmSysDate = dtResult.Rows(0).Item("SysDate")                 'サーバー日付
            Else
                '取得できなかったときはエラー
                puErrMsg = M0201_E0000
                MsgBox(puErrMsg, MsgBoxStyle.Critical)
                Return False
            End If

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = ex.Message
            Return False
        Finally
            dtResult.Dispose()
            Adapter.Dispose()
        End Try
    End Function


    ''' <summary>利用者情報更新用SQL作成処理
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>入力内容を利用者マスタに反映する
    ''' <para>作成情報：2015/08/25 ozawa
    ''' <p>改訂情報：</p>
    ''' </para>
    ''' </remarks>
    Public Function UpdateNewRiyoshaData(ByRef Tsx As NpgsqlTransaction, _
                                  ByVal Cn As NpgsqlConnection, _
                                  ByRef dataEXTM0202 As DataEXTM0202) As Boolean
        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cmd As New NpgsqlCommand            'SQLコマンド

        Try
            '利用者情報更新用SQL
            If SqlEXTM0202.UpdateRiyoshaData(Cmd, Cn, dataEXTM0202) = False Then
                Return False
            End If

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "利用者情報更新", Nothing, Cmd)

            'SQL実行
            Cmd.ExecuteNonQuery()

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'ロールバック
            If Tsx IsNot Nothing Then
                Tsx.Rollback()
            End If
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = M0201_E0000 & ex.Message
            Return False
        Finally
            Cmd.Dispose()
        End Try

    End Function

    ''' <summary>新規利用者番号、システム日付取得
    ''' </summary>
    ''' <param name="Cn"></param>
    ''' <param name="dataEXTM0202"></param>
    ''' <returns></returns>
    ''' <remarks>新規利用者番号、システム日付を取得する
    ''' <para>作成情報：2015/08/25 ozawa
    ''' <p>改訂情報：</p>
    ''' </para>
    ''' </remarks>
    Public Function GetNewRiyoshaCd(ByVal Cn As NpgsqlConnection, ByRef dataEXTM0202 As DataEXTM0202) As Boolean
        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Adapter As New NpgsqlDataAdapter
        Dim dtResult As New DataTable

        Try
            '新規番号取得（SELECT）用SQLを作成
            If SqlEXTM0202.GetNewRiyoCd(Adapter, Cn, dataEXTM0202) = False Then
                Return False
            End If

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "新規利用者番号、システム日付取得", Nothing, Adapter.SelectCommand)

            'SQL実行
            Adapter.Fill(dtResult)

            'データが取得できた場合、データクラスに取得データをセット
            If dtResult.Rows.Count > 0 Then
                'dataEXTM0202.PropNewRiyo_cd = "U" & dtResult.Rows(0).Item("Riyosha_cd")   '新規利用者番号     2015.10.06 UPDATE h.hagiwara 利用者CD上1文字目の値変更
                dataEXTM0202.PropNewRiyo_cd = "C" & dtResult.Rows(0).Item("Riyosha_cd")    '新規利用者番号     2015.10.06 UPDATE h.hagiwara 利用者CD上1文字目の値変更
                dataEXTM0202.PropDtmSysDate = dtResult.Rows(0).Item("SysDate")                 'サーバー日付
            Else
                '取得できなかったときはエラー
                puErrMsg = M0201_E0000
                MsgBox(puErrMsg, MsgBoxStyle.Critical)
                Return False
            End If

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = ex.Message
            Return False
        Finally
            dtResult.Dispose()
            Adapter.Dispose()
        End Try

    End Function

    ''' <summary>利用者情報新規登録SQL作成処理
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>入力内容を利用者マスタに新規登録する
    ''' <para>作成情報：2015/08/25 ozawa
    ''' <p>作成情報：</p>
    ''' </para>
    ''' </remarks>
    Public Function InsertNewRiyoshaData(ByRef Tsx As NpgsqlTransaction, _
                                  ByVal Cn As NpgsqlConnection, _
                                  ByVal dataEXTM0202 As DataEXTM0202) As Boolean
        '開始ログ出力
        CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "START", Nothing, Nothing)

        '変数宣言
        Dim Cmd As New NpgsqlCommand            'SQLコマンド

        Try
            '利用者情報新規登録用SQL
            If SqlEXTM0202.InsertRiyoshaData(Cmd, Cn, dataEXTM0202) = False Then
                Return False
            End If

            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.DEBUG_Lv, "利用者情報登録", Nothing, Cmd)

            'SQL実行
            Cmd.ExecuteNonQuery()

            '終了ログ出力
            CommonLogic.WriteLog(Common.LogLevel.TRACE_Lv, "END", Nothing, Nothing)

            '正常処理終了
            Return True

        Catch ex As Exception
            'ログ出力
            CommonLogic.WriteLog(Common.LogLevel.ERROR_Lv, ex.Message, Nothing, Nothing)
            'ロールバック
            If Tsx IsNot Nothing Then
                Tsx.Rollback()
            End If
            'メッセージ変数にエラーメッセージを格納
            puErrMsg = ex.Message
            Return False
        Finally
            Cmd.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' 全角カナチェック
    ''' </summary>
    ''' <param name="argTxt">[IN]チェック対象文字列</param>
    ''' <returns>Boolean  チェック結果    true  ＯＫ  false  ＮＧ</returns>
    ''' <remarks>対象文字列が半角カナ文字のみかチェックする
    ''' <para>作成情報：2015/08/25 ozawa
    ''' <p>改定情報：</p>
    ''' </para></remarks>
    Public Function IsFullKana(ByVal argTxt As String) As Boolean

        Dim i As Integer
        Dim length As Integer

        length = Len(argTxt)

        If length = 0 Then
            Return True
        End If

        For i = 1 To length
            If Not Mid(argTxt, i, 1) Like "^[ア-ン]*$" = False Then
                Return False
            End If
        Next i

        Return True

    End Function

End Class
