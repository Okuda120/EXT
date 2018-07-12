Public Class DataEXTA0101

    'パラメータ変数宣言(検索条件)
    Private ppTxtUserId As TextBox                  'ユーザーID
    Private ppTxtPassword As TextBox                'パスワード
    Private ppStrUserName As String                 'ユーザー名
    Private ppStrMailAddr As String                     'Mail
    Private ppStrCode_BUSHO As String        'ペーパー版資材発送フラグ
    Private ppStrFlg_SHONIN As String      'ペーパー版成績処理フラグ
    Private ppStrFlg_MST As String      '学校版成績処理フラグ

    ''' <summary>
    ''' プロパティセット【ユーザーID】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppTxtUserId</returns>
    ''' <remarks><para>作成情報：2015/08/04 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropTxtUserId()
        Get
            Return ppTxtUserId
        End Get
        Set(ByVal value)
            ppTxtUserId = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【パスワード】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppTxtPassword</returns>
    ''' <remarks><para>作成情報：2015/08/04 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropTxtPassword()
        Get
            Return ppTxtPassword
        End Get
        Set(ByVal value)
            ppTxtPassword = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【ユーザー名】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrUserName</returns>
    ''' <remarks><para>作成情報：2015/08/04 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrUserName()
        Get
            Return ppStrUserName
        End Get
        Set(ByVal value)
            ppStrUserName = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【メールアドレス】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrMailAddr</returns>
    ''' <remarks><para>作成情報：2015/08/04 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrMailAddr()
        Get
            Return ppStrMailAddr
        End Get
        Set(ByVal value)
            ppStrMailAddr = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【部署コード】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrCode_BUSHO</returns>
    ''' <remarks><para>作成情報：2015/08/04 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrCode_BUSHO()
        Get
            Return ppStrCode_BUSHO
        End Get
        Set(ByVal value)
            ppStrCode_BUSHO = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【承認フラグ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrFlg_SHONIN</returns>
    ''' <remarks><para>作成情報：2015/08/04 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrFlg_SHOHIN()
        Get
            Return ppStrFlg_SHONIN
        End Get
        Set(ByVal value)
            ppStrFlg_SHONIN = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【マスタフラグ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrFlg_MST</returns>
    ''' <remarks><para>作成情報：2015/08/04 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrFlg_MST()
        Get
            Return ppStrFlg_MST
        End Get
        Set(ByVal value)
            ppStrFlg_MST = value
        End Set
    End Property

End Class
