Public Class DataEXTZ0103

    'パラメータ変数宣言
    Private ppStrShisetsuKbn As String          '施設区分
    Private ppStrShiyoNenFrom As String         '使用年（From）
    Private ppStrShiyoTsukiFrom As String       '使用月（From）
    Private ppStrShiyoNenTo As String           '使用年（To）
    Private ppStrShiyoTsukiTo As String         '使用月（To）
    Private ppStrRiyoNm As String               '利用者名
    Private ppStrRiyoNmKana As String           '利用者名カナ
    Private ppvwDayUriageTheatre As FarPoint.Win.Spread.FpSpread    '日別売上一覧（シアター）
    Private ppvwDayUriageStudio As FarPoint.Win.Spread.FpSpread     '日別売上一覧（スタジオ）
    Private ppvwSeikyuChoseiTheatre As FarPoint.Win.Spread.FpSpread '請求時調整額（シアター）
    Private ppvwSeikyuChoseiStudio As FarPoint.Win.Spread.FpSpread  '請求時調整額（スタジオ）
    Private ppDtDayUriage As DataTable                              '日別売上データ
    Private ppDtSeikyuChosei As DataTable                           '請求時調整額データ

    ''' <summary>
    ''' プロパティセット【施設区分】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrShisetsuKbn</returns>
    ''' <remarks><para>作成情報：2015/09/16 h.endo
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
    ''' プロパティセット【使用年（From）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrShiyoNenFrom</returns>
    ''' <remarks><para>作成情報：2015/09/16 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrShiyoNenFrom() As String
        Get
            Return ppStrShiyoNenFrom
        End Get
        Set(ByVal value As String)
            ppStrShiyoNenFrom = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【使用月（From）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrShiyoTsukiFrom</returns>
    ''' <remarks><para>作成情報：2015/09/16 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrShiyoTsukiFrom() As String
        Get
            Return ppStrShiyoTsukiFrom
        End Get
        Set(ByVal value As String)
            ppStrShiyoTsukiFrom = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【使用年（To）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrShiyoNenTo</returns>
    ''' <remarks><para>作成情報：2015/09/16 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrShiyoNenTo() As String
        Get
            Return ppStrShiyoNenTo
        End Get
        Set(ByVal value As String)
            ppStrShiyoNenTo = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【使用月（To）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrShiyoTsukiTo</returns>
    ''' <remarks><para>作成情報：2015/09/16 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrShiyoTsukiTo() As String
        Get
            Return ppStrShiyoTsukiTo
        End Get
        Set(ByVal value As String)
            ppStrShiyoTsukiTo = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者名】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyoNm</returns>
    ''' <remarks><para>作成情報：2015/09/16 h.endo
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
    ''' プロパティセット【利用者名カナ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyoNmKana</returns>
    ''' <remarks><para>作成情報：2015/09/16 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrRiyoNmKana() As String
        Get
            Return ppStrRiyoNmKana
        End Get
        Set(ByVal value As String)
            ppStrRiyoNmKana = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【日別売上一覧（シアター）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppvwDayUriageTheatre</returns>
    ''' <remarks><para>作成情報：2015/09/16 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropvwDayUriageTheatre() As FarPoint.Win.Spread.FpSpread
        Get
            Return ppvwDayUriageTheatre
        End Get
        Set(ByVal value As FarPoint.Win.Spread.FpSpread)
            ppvwDayUriageTheatre = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【日別売上一覧（スタジオ）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppvwDayUriageStudio</returns>
    ''' <remarks><para>作成情報：2015/09/16 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropvwDayUriageStudio() As FarPoint.Win.Spread.FpSpread
        Get
            Return ppvwDayUriageStudio
        End Get
        Set(ByVal value As FarPoint.Win.Spread.FpSpread)
            ppvwDayUriageStudio = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【請求時調整額（シアター）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppvwSeikyuChoseiTheatre</returns>
    ''' <remarks><para>作成情報：2015/09/16 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropvwSeikyuChoseiTheatre() As FarPoint.Win.Spread.FpSpread
        Get
            Return ppvwSeikyuChoseiTheatre
        End Get
        Set(ByVal value As FarPoint.Win.Spread.FpSpread)
            ppvwSeikyuChoseiTheatre = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【請求時調整額（スタジオ）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppvwSeikyuChoseiStudio</returns>
    ''' <remarks><para>作成情報：2015/09/16 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropvwSeikyuChoseiStudio() As FarPoint.Win.Spread.FpSpread
        Get
            Return ppvwSeikyuChoseiStudio
        End Get
        Set(ByVal value As FarPoint.Win.Spread.FpSpread)
            ppvwSeikyuChoseiStudio = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【日別売上データ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtDayUriage</returns>
    ''' <remarks><para>作成情報：2015/09/16 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtDayUriage() As DataTable
        Get
            Return ppDtDayUriage
        End Get
        Set(ByVal value As DataTable)
            ppDtDayUriage = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【請求調整データ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtSeikyuChosei</returns>
    ''' <remarks><para>作成情報：2015/09/16 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtSeikyuChosei() As DataTable
        Get
            Return ppDtSeikyuChosei
        End Get
        Set(ByVal value As DataTable)
            ppDtSeikyuChosei = value
        End Set
    End Property

End Class
