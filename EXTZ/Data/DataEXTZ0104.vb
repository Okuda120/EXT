Public Class DataEXTZ0104

    'パラメータ変数宣言
    Private ppStrShisetsuKbn As String          '施設区分
    Private ppStrKikanNen As String             '期間年
    Private ppStrKikanTsuki As String           '期間月
    Private ppvwRiyoJokyoTheatre As FarPoint.Win.Spread.FpSpread    '利用状況一覧（シアター）
    Private ppvwRiyoJokyoStudio As FarPoint.Win.Spread.FpSpread     '利用状況一覧（スタジオ）
    Private ppDtRiyoJokyo1 As DataTable                             '利用状況データ１
    Private ppDtRiyoJokyo2 As DataTable                             '利用状況データ２
    Private ppDtRiyoJokyo3 As DataTable                             '利用状況データ３
    Private ppDtRiyoJokyo4 As DataTable                             '利用状況データ４
    Private ppDtRiyoJokyo5 As DataTable                             '利用状況データ５
    Private ppDtRiyoJokyo6 As DataTable                             '利用状況データ６
    Private ppDtRiyoJokyo7 As DataTable                             '利用状況データ７
    Private ppDtRiyoJokyo8 As DataTable                             '利用状況データ８
    Private ppDtRiyoJokyo9 As DataTable                             '利用状況データ９
    Private ppDtRiyoJokyo10 As DataTable                            '利用状況データ１０
    Private ppDtRiyoJokyo11 As DataTable                            '利用状況データ１１
    Private ppDtRiyoJokyo12 As DataTable                            '利用状況データ１２

    ''' <summary>
    ''' プロパティセット【施設区分】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrShisetsuKbn</returns>
    ''' <remarks><para>作成情報：2015/09/24 h.endo
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
    ''' プロパティセット【期間年】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrKikanNen</returns>
    ''' <remarks><para>作成情報：2015/09/24 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrKikanNen() As String
        Get
            Return ppStrKikanNen
        End Get
        Set(ByVal value As String)
            ppStrKikanNen = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【期間月】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrKikanTsuki</returns>
    ''' <remarks><para>作成情報：2015/09/24 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrKikanTsuki() As String
        Get
            Return ppStrKikanTsuki
        End Get
        Set(ByVal value As String)
            ppStrKikanTsuki = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用状況一覧（シアター）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppvwRiyoJokyoTheatre</returns>
    ''' <remarks><para>作成情報：2015/09/24 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropvwRiyoJokyoTheatre() As FarPoint.Win.Spread.FpSpread
        Get
            Return ppvwRiyoJokyoTheatre
        End Get
        Set(ByVal value As FarPoint.Win.Spread.FpSpread)
            ppvwRiyoJokyoTheatre = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用状況一覧（スタジオ）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppvwRiyoJokyoStudio</returns>
    ''' <remarks><para>作成情報：2015/09/24 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropvwRiyoJokyoStudio() As FarPoint.Win.Spread.FpSpread
        Get
            Return ppvwRiyoJokyoStudio
        End Get
        Set(ByVal value As FarPoint.Win.Spread.FpSpread)
            ppvwRiyoJokyoStudio = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用状況データ１】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtRiyoJokyo1</returns>
    ''' <remarks><para>作成情報：2015/09/24 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtRiyoJokyo1() As DataTable
        Get
            Return ppDtRiyoJokyo1
        End Get
        Set(ByVal value As DataTable)
            ppDtRiyoJokyo1 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用状況データ２】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtRiyoJokyo2</returns>
    ''' <remarks><para>作成情報：2015/09/24 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtRiyoJokyo2() As DataTable
        Get
            Return ppDtRiyoJokyo2
        End Get
        Set(ByVal value As DataTable)
            ppDtRiyoJokyo2 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用状況データ３】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtRiyoJokyo3</returns>
    ''' <remarks><para>作成情報：2015/09/24 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtRiyoJokyo3() As DataTable
        Get
            Return ppDtRiyoJokyo3
        End Get
        Set(ByVal value As DataTable)
            ppDtRiyoJokyo3 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用状況データ４】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtRiyoJokyo4</returns>
    ''' <remarks><para>作成情報：2015/09/24 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtRiyoJokyo4() As DataTable
        Get
            Return ppDtRiyoJokyo4
        End Get
        Set(ByVal value As DataTable)
            ppDtRiyoJokyo4 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用状況データ５】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtRiyoJokyo5</returns>
    ''' <remarks><para>作成情報：2015/09/24 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtRiyoJokyo5() As DataTable
        Get
            Return ppDtRiyoJokyo5
        End Get
        Set(ByVal value As DataTable)
            ppDtRiyoJokyo5 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用状況データ６】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtRiyoJokyo6</returns>
    ''' <remarks><para>作成情報：2015/09/24 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtRiyoJokyo6() As DataTable
        Get
            Return ppDtRiyoJokyo6
        End Get
        Set(ByVal value As DataTable)
            ppDtRiyoJokyo6 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用状況データ７】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtRiyoJokyo7</returns>
    ''' <remarks><para>作成情報：2015/09/24 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtRiyoJokyo7() As DataTable
        Get
            Return ppDtRiyoJokyo7
        End Get
        Set(ByVal value As DataTable)
            ppDtRiyoJokyo7 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用状況データ８】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtRiyoJokyo8</returns>
    ''' <remarks><para>作成情報：2015/09/24 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtRiyoJokyo8() As DataTable
        Get
            Return ppDtRiyoJokyo8
        End Get
        Set(ByVal value As DataTable)
            ppDtRiyoJokyo8 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用状況データ９】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtRiyoJokyo9</returns>
    ''' <remarks><para>作成情報：2015/09/24 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtRiyoJokyo9() As DataTable
        Get
            Return ppDtRiyoJokyo9
        End Get
        Set(ByVal value As DataTable)
            ppDtRiyoJokyo9 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用状況データ１０】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtRiyoJokyo10</returns>
    ''' <remarks><para>作成情報：2015/09/24 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtRiyoJokyo10() As DataTable
        Get
            Return ppDtRiyoJokyo10
        End Get
        Set(ByVal value As DataTable)
            ppDtRiyoJokyo10 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用状況データ１１】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtRiyoJokyo11</returns>
    ''' <remarks><para>作成情報：2015/09/24 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtRiyoJokyo11() As DataTable
        Get
            Return ppDtRiyoJokyo11
        End Get
        Set(ByVal value As DataTable)
            ppDtRiyoJokyo11 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用状況データ１２】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtRiyoJokyo12</returns>
    ''' <remarks><para>作成情報：2015/09/24 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtRiyoJokyo12() As DataTable
        Get
            Return ppDtRiyoJokyo12
        End Get
        Set(ByVal value As DataTable)
            ppDtRiyoJokyo12 = value
        End Set
    End Property


End Class
