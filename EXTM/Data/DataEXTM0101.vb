Imports FarPoint.Win.Spread

Public Class DataEXTM0101
    'フォームオブジェクト
    Private ppTxtSearchUserId As TextBox            '検索ユーザＩＤ
    Private ppTxtSearchKanjiName As TextBox         '検索漢字氏名
    Private ppTxtSearchMail As TextBox              '検索メールアドレス
    Private ppCmbSearchBushoName As ComboBox        '検索部署名
    Private ppChkShoninFlg As CheckBox              '承認者権限フラグ
    Private ppChkStsFlg As CheckBox                 '無効フラグ
    Private ppChkMstFlg As CheckBox                 'マスタ操作権限フラグ
    Private ppBtnSearch As Button                   '検索ボタン
    Private ppVwList As FpSpread                    '一覧シート

    'データ
    Private ppDtExtUsrMasta As DataTable            '検索結果を格納するデータテーブル
    Private ppDtExtBushoMasta As DataTable          '部署名を格納するデータテーブル

    'フッタ
    Private ppBtnReg As Button                      'フッタ：登録ボタン
    Private ppBtnBack As Button                     'フッタ：戻るボタン


    ''' <summary>
    ''' プロパティセット【検索ユーザID】
    ''' </summary>
    ''' <value></value>
    ''' <returns> ppTxtSearchUserId</returns>
    ''' <remarks><para>作成情報：2015/08/11 hayabuchi
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropTxtSearchUserID() As TextBox
        Get
            Return ppTxtSearchUserId
        End Get
        Set(ByVal value As TextBox)
            ppTxtSearchUserId = value
        End Set
    End Property
    ''' <summary>
    ''' プロパティセット【検索漢字氏名】
    ''' </summary>
    ''' <value></value>
    ''' <returns> ppTxtSearchKanjiName</returns>
    ''' <remarks><para>作成情報：2015/08/11 hayabuchi
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropTxtSearchKanjiName() As TextBox
        Get
            Return ppTxtSearchKanjiName
        End Get
        Set(ByVal value As TextBox)
            ppTxtSearchKanjiName = value
        End Set
    End Property
    ''' <summary>
    ''' プロパティセット【検索メールアドレス】
    ''' </summary>
    ''' <value></value>
    ''' <returns> ppTxtSearchMail</returns>
    ''' <remarks><para>作成情報：2015/08/11 hayabuchi
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropTxtSearchMail() As TextBox
        Get
            Return ppTxtSearchMail
        End Get
        Set(ByVal value As TextBox)
            ppTxtSearchMail = value
        End Set
    End Property
    ''' <summary>
    ''' プロパティセット【検索部署名】
    ''' </summary>
    ''' <value></value>
    ''' <returns> ppCmbSearchBushoName</returns>
    ''' <remarks><para>作成情報：2015/08/11 hayabuchi
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropCmbSearchBushoName() As ComboBox
        Get
            Return ppCmbSearchBushoName
        End Get
        Set(ByVal value As ComboBox)
            ppCmbSearchBushoName = value
        End Set
    End Property
    ''' <summary>
    ''' プロパティセット【検索承認者権限フラグ】
    ''' </summary>
    ''' <value></value>
    ''' <returns> ppChkShoninFlg</returns>
    ''' <remarks><para>作成情報：2015/08/11 hayabuchi
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropChkShoninFlg() As CheckBox
        Get
            Return ppChkShoninFlg
        End Get
        Set(ByVal value As CheckBox)
            ppChkShoninFlg = value
        End Set
    End Property
    ''' <summary>
    ''' プロパティセット【検索マスタ操作権限】
    ''' </summary>
    ''' <value></value>
    ''' <returns> ppChkMstFlg</returns>
    ''' <remarks><para>作成情報：2015/08/11 hayabuchi
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropChkMstFlg() As CheckBox
        Get
            Return ppChkMstFlg
        End Get
        Set(ByVal value As CheckBox)
            ppChkMstFlg = value
        End Set
    End Property
    ''' <summary>
    ''' プロパティセット【無効フラグ】
    ''' </summary>
    ''' <value></value>
    ''' <returns> ppChkStsFlg</returns>
    ''' <remarks><para>作成情報：2015/08/11 hayabuchi
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropChkStsFlg() As CheckBox
        Get
            Return ppChkStsFlg
        End Get
        Set(ByVal value As CheckBox)
            ppChkStsFlg = value
        End Set
    End Property
    ''' <summary>
    ''' プロパティセット【データテーブル】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtExtUsrMasta</returns>
    ''' <remarks><para>作成情報：2015/08/11 hayabuchi
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtExtUsrMasta() As DataTable
        Get
            Return ppDtExtUsrMasta
        End Get
        Set(ByVal value As DataTable)
            ppDtExtUsrMasta = value
        End Set
    End Property
    ''' <summary>
    ''' プロパティセット【検索ボタン】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppBtnSearch</returns>
    ''' <remarks><para>作成情報：2015/08/11 hayabuchi
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropBtnSearch() As Button
        Get
            Return ppBtnSearch
        End Get
        Set(ByVal value As Button)
            ppBtnSearch = value
        End Set
    End Property
    ''' <summary>
    ''' プロパティセット【一覧シート】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppVwList</returns>
    ''' <remarks><para>作成情報：2015/08/11 hayabuchi
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropVwList() As FpSpread
        Get
            Return ppVwList
        End Get
        Set(ByVal value As FpSpread)
            ppVwList = value
        End Set
    End Property
    ''' <summary>
    ''' プロパティセット【部署名を格納するデータテーブル】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtExtBushoMasta</returns>
    ''' <remarks><para>作成情報：2015/08/11 hayabuchi
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtExtBushoMasta() As DataTable
        Get
            Return ppDtExtBushoMasta
        End Get
        Set(ByVal value As DataTable)
            ppDtExtBushoMasta = value
        End Set
    End Property
    ''' <summary>
    ''' プロパティセット【フッタ：登録ボタン】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppBtnReg</returns>
    ''' <remarks><para>作成情報：2015/08/11 hayabuchi
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropBtnReg() As Button
        Get
            Return ppBtnReg
        End Get
        Set(ByVal value As Button)
            ppBtnReg = value
        End Set
    End Property
    ''' <summary>
    ''' プロパティセット【フッタ：戻るボタン】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppBtnBack</returns>
    ''' <remarks><para>作成情報：2015/08/11 hayabuchi
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropBtnBack() As Button
        Get
            Return ppBtnBack
        End Get
        Set(ByVal value As Button)
            ppBtnBack = value
        End Set
    End Property
End Class
