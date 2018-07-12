Public Class DataEXTZ0102

    'パラメータ変数宣言
    Private ppStrShisetsuKbn As String          '施設区分
    Private ppStrSeikyuDtFrom As String         '請求日（From）
    Private ppStrSeikyuDtTo As String           '請求日（To）
    Private ppStrNyukinYoteiDtFrom As String    '入金予定日（From）
    Private ppStrNyukinYoteiDtTo As String      '入金予定日（To）
    Private ppStrNyukinDtFrom As String         '入金日（From）
    Private ppStrNyukinDtTo As String           '入金日（To）
    Private ppStrAiteNm As String               '相手先名
    Private ppStrAiteNmKana As String           '相手先名カナ
    Private ppStrRiyoNm As String               '利用者名
    Private ppStrSaijiNm As String              '催事名
    Private ppStrSeikyuIraiNo As String         '請求依頼番号
    Private ppBlnMinyukin As Boolean            '未入金フラグ
    Private ppvwSeikyuTheatre As FarPoint.Win.Spread.FpSpread       '請求一覧（シアター）
    Private ppvwSeikyuStudio As FarPoint.Win.Spread.FpSpread        '請求一覧（スタジオ）
    Private ppDtSeikyu As DataTable             '請求データ

    ''' <summary>
    ''' プロパティセット【施設区分】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrShisetsuKbn</returns>
    ''' <remarks><para>作成情報：2015/09/15 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrShisetsuKbn() As String
        Get
            Return ppStrShisetsuKbn
        End Get
        Set(ByVal value As String)
            ppStrShisetsuKbn = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【請求日（From）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrSeikyuDtFrom</returns>
    ''' <remarks><para>作成情報：2015/09/15 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrSeikyuDtFrom() As String
        Get
            Return ppStrSeikyuDtFrom
        End Get
        Set(ByVal value As String)
            ppStrSeikyuDtFrom = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【請求日（To）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrSeikyuDtTo</returns>
    ''' <remarks><para>作成情報：2015/09/15 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrSeikyuDtTo() As String
        Get
            Return ppStrSeikyuDtTo
        End Get
        Set(ByVal value As String)
            ppStrSeikyuDtTo = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【入金予定日（From）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrNyukinYoteiDtFrom</returns>
    ''' <remarks><para>作成情報：2015/09/15 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrNyukinYoteiDtFrom() As String
        Get
            Return ppStrNyukinYoteiDtFrom
        End Get
        Set(ByVal value As String)
            ppStrNyukinYoteiDtFrom = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【入金予定日（To）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrSeikyuDtTo</returns>
    ''' <remarks><para>作成情報：2015/09/15 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrNyukinYoteiDtTo() As String
        Get
            Return ppStrNyukinYoteiDtTo
        End Get
        Set(ByVal value As String)
            ppStrNyukinYoteiDtTo = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【入金日（From）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrNyukinDtFrom</returns>
    ''' <remarks><para>作成情報：2015/09/15 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrNyukinDtFrom() As String
        Get
            Return ppStrNyukinDtFrom
        End Get
        Set(ByVal value As String)
            ppStrNyukinDtFrom = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【入金日（To）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrNyukinDtTo</returns>
    ''' <remarks><para>作成情報：2015/09/15 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrNyukinDtTo() As String
        Get
            Return ppStrNyukinDtTo
        End Get
        Set(ByVal value As String)
            ppStrNyukinDtTo = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【相手先名】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrAiteNm</returns>
    ''' <remarks><para>作成情報：2015/09/15 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrAiteNm() As String
        Get
            Return ppStrAiteNm
        End Get
        Set(ByVal value As String)
            ppStrAiteNm = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【相手先カナ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrAiteNmKana</returns>
    ''' <remarks><para>作成情報：2015/09/15 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrAiteNmKana() As String
        Get
            Return ppStrAiteNmKana
        End Get
        Set(ByVal value As String)
            ppStrAiteNmKana = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者名】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyoNm</returns>
    ''' <remarks><para>作成情報：2015/09/15 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrRiyoNm() As String
        Get
            Return ppStrRiyoNm
        End Get
        Set(ByVal value As String)
            ppStrRiyoNm = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【催事名】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrSaijiNm</returns>
    ''' <remarks><para>作成情報：2015/09/15 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrSaijiNm() As String
        Get
            Return ppStrSaijiNm
        End Get
        Set(ByVal value As String)
            ppStrSaijiNm = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【請求依頼番号】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrSeikyuIraiNo</returns>
    ''' <remarks><para>作成情報：2015/09/15 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrSeikyuIraiNo() As String
        Get
            Return ppStrSeikyuIraiNo
        End Get
        Set(ByVal value As String)
            ppStrSeikyuIraiNo = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【未入金フラグ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppBlnMinyukin</returns>
    ''' <remarks><para>作成情報：2015/09/15 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropBlnMinyukin() As Boolean
        Get
            Return ppBlnMinyukin
        End Get
        Set(ByVal value As Boolean)
            ppBlnMinyukin = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【請求一覧（シアター）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppvwSeikyuTheatre</returns>
    ''' <remarks><para>作成情報：2015/09/15 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropvwSeikyuTheatre() As FarPoint.Win.Spread.FpSpread
        Get
            Return ppvwSeikyuTheatre
        End Get
        Set(ByVal value As FarPoint.Win.Spread.FpSpread)
            ppvwSeikyuTheatre = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【請求一覧（スタジオ）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppvwSeikyuStudio</returns>
    ''' <remarks><para>作成情報：2015/09/15 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropvwSeikyuStudio() As FarPoint.Win.Spread.FpSpread
        Get
            Return ppvwSeikyuStudio
        End Get
        Set(ByVal value As FarPoint.Win.Spread.FpSpread)
            ppvwSeikyuStudio = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【請求データ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtSeikyu</returns>
    ''' <remarks><para>作成情報：2015/09/15 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtSeikyu() As DataTable
        Get
            Return ppDtSeikyu
        End Get
        Set(ByVal value As DataTable)
            ppDtSeikyu = value
        End Set
    End Property


End Class
