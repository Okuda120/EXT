Imports FarPoint.Win.Spread

Public Class DataEXTZ0205

    Private ppBlnChangeFlg As Boolean
    Private ppStrPrjCd As String
    Private ppStrPrjNm As String
    Private ppUchiCd As String
    Private ppUchiNm As String
    Private ppStrRiyobi As String
    Private ppDtResult As DataTable
    Private ppStrResPrjCd As String
    Private ppStrResPrjNm As String
    Private ppStrResUchiCd As String
    Private ppStrResUchiNm As String
    Private fbResult As FpSpread          'スプレッドシート(2015/11/30 hayabuchi チェック制御)
    Private ppIntCheckRow As Integer      'チェック行番号(2015/11/30 hayabuchi チェック制御)

    ''' <summary>
    ''' プロパティセット【変更フラグ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppBlnChangeFlg</returns>
    ''' <remarks><para>作成情報：2015/09/01 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropBlnChangeFlg()
        Get
            Return ppBlnChangeFlg
        End Get
        Set(ByVal value)
            ppBlnChangeFlg = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【検索キー】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrPrjCd</returns>
    ''' <remarks><para>作成情報：2015/09/01 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrPrjCd()
        Get
            Return ppStrPrjCd
        End Get
        Set(ByVal value)
            ppStrPrjCd = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【検索キー】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrPrjNm</returns>
    ''' <remarks><para>作成情報：2015/09/01 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrPrjNm()
        Get
            Return ppStrPrjNm
        End Get
        Set(ByVal value)
            ppStrPrjNm = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【検索キー】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppUchiCd</returns>
    ''' <remarks><para>作成情報：2015/09/01 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropUchiCd()
        Get
            Return ppUchiCd
        End Get
        Set(ByVal value)
            ppUchiCd = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【検索キー】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppUchiNm</returns>
    ''' <remarks><para>作成情報：2015/09/01 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropUchiNm()
        Get
            Return ppUchiNm
        End Get
        Set(ByVal value)
            ppUchiNm = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【検索キー】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyobi</returns>
    ''' <remarks><para>作成情報：2015/09/01 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrRiyobi()
        Get
            Return ppStrRiyobi
        End Get
        Set(ByVal value)
            ppStrRiyobi = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【検索結果】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtResult</returns>
    ''' <remarks><para>作成情報：2015/09/01 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtResult()
        Get
            Return ppDtResult
        End Get
        Set(ByVal value)
            ppDtResult = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【選択値】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrResPrjCd</returns>
    ''' <remarks><para>作成情報：2015/09/01 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrResPrjCd()
        Get
            Return ppStrResPrjCd
        End Get
        Set(ByVal value)
            ppStrResPrjCd = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【選択値】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrResPrjNm</returns>
    ''' <remarks><para>作成情報：2015/09/01 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrResPrjNm()
        Get
            Return ppStrResPrjNm
        End Get
        Set(ByVal value)
            ppStrResPrjNm = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【選択値】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrResUchiCd</returns>
    ''' <remarks><para>作成情報：2015/09/01 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrResUchiCd()
        Get
            Return ppStrResUchiCd
        End Get
        Set(ByVal value)
            ppStrResUchiCd = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【選択値】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrResUchiNm</returns>
    ''' <remarks><para>作成情報：2015/09/01 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrResUchiNm()
        Get
            Return ppStrResUchiNm
        End Get
        Set(ByVal value)
            ppStrResUchiNm = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【スプレッドシート】
    ''' </summary>
    ''' <value></value>
    ''' <returns>fbResult</returns>
    ''' <remarks><para>作成情報：2015/11/30 hayabuchi
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropResult() As FpSpread
        Get
            Return fbResult
        End Get
        Set(ByVal value As FpSpread)
            fbResult = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【チェック行番号】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppIntCheckRow</returns>
    ''' <remarks><para>作成情報：2015/11/30 hayabuchi
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropIntCheckRow() As Integer
        Get
            Return ppIntCheckRow
        End Get
        Set(ByVal value As Integer)
            ppIntCheckRow = value
        End Set
    End Property
End Class
