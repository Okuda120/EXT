Public Class DataEXTB0101

    'パラメータ変数宣言
    Private ppTxtYear As TextBox    '年
    Private ppTxtMonth As TextBox   '月
    Private ppStrYear As String    '検索条件・年
    Private ppStrMonth As String   '検索条件・月
    Private ppStrYYMMDD As String  '検索条件・年月日
    Private ppvwCalandarFirst As FarPoint.Win.Spread.FpSpread       'カレンダー一覧（前半）
    Private ppvwCalandarSecond As FarPoint.Win.Spread.FpSpread      'カレンダー一覧（後半）
    Private ppvwCancelWait As FarPoint.Win.Spread.FpSpread          'キャンセル待ち一覧
    Private ppvwCancelDone As FarPoint.Win.Spread.FpSpread          'キャンセル済一覧
    Private ppDtShukusaijitsu As DataTable                          '祝祭日データ
    Private ppDtYoyakuList As DataTable                             '予約一覧データ
    Private ppDtMikakuYoyakuList As DataTable                       '予約未確定一覧データ
    Private ppDtDateKakuOnlyCancelWaitList As DataTable             'キャンセル待ちデータ（日付確定のみ）
    Private ppDtAllCancelWaitList As DataTable                      'キャンセル待ちデータ（全て）
    Private ppDtMikakuCancelWaitList As DataTable                   'キャンセル待ち未確定データ
    Private ppDtCancelDoneList As DataTable                         'キャンセル済データ
    Private ppDtEtcUseYoyakuList As DataTable                       'その他利用予約一覧データ
    Private ppDtKyuykanMaiteList As DataTable                       '休館メンテ一覧データ
    Private ppDicYoyakuInfo As Dictionary(Of String, String)        '予約情報（カレンダーからのクリック対応に基づく紐付け用）
    Private ppDicCancelInfo As Dictionary(Of String, String)        'キャンセル情報（カレンダーからのクリック対応に基づく紐付け用）

    ''' <summary>
    ''' プロパティセット【年】
    ''' </summary>
    ''' <value></value>
    ''' <returns>TxtYear</returns>
    ''' <remarks><para>作成情報：2015/08/07 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropTxtYear() As TextBox
        Get
            Return ppTxtYear
        End Get
        Set(ByVal value As TextBox)
            ppTxtYear = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【月】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppTxtMonth</returns>
    ''' <remarks><para>作成情報：2015/08/07 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropTxtMonth() As TextBox
        Get
            Return ppTxtMonth
        End Get
        Set(ByVal value As TextBox)
            ppTxtMonth = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【検索条件・年】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrYear</returns>
    ''' <remarks><para>作成情報：2015/08/07 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrYear() As String
        Get
            Return ppStrYear
        End Get
        Set(ByVal value As String)
            ppStrYear = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【検索条件・月】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrMonth</returns>
    ''' <remarks><para>作成情報：2015/08/07 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrMonth() As String
        Get
            Return ppStrMonth
        End Get
        Set(ByVal value As String)
            ppStrMonth = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【検索条件・年月日】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrYYMMDD</returns>
    ''' <remarks><para>作成情報：2015/08/18 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrYYMMDD() As String
        Get
            Return ppStrYYMMDD
        End Get
        Set(ByVal value As String)
            ppStrYYMMDD = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【カレンダー一覧（前半）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppvwCalandarFirst</returns>
    ''' <remarks><para>作成情報：2015/08/07 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropvwCalandarFirst() As FarPoint.Win.Spread.FpSpread
        Get
            Return ppvwCalandarFirst
        End Get
        Set(ByVal value As FarPoint.Win.Spread.FpSpread)
            ppvwCalandarFirst = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【カレンダー一覧（後半）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppvwCalandarSecond</returns>
    ''' <remarks><para>作成情報：2015/08/07 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropvwCalandarSecond() As FarPoint.Win.Spread.FpSpread
        Get
            Return ppvwCalandarSecond
        End Get
        Set(ByVal value As FarPoint.Win.Spread.FpSpread)
            ppvwCalandarSecond = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【キャンセル待ち一覧】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppvwCancelWait</returns>
    ''' <remarks><para>作成情報：2015/08/07 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropvwCancelWait() As FarPoint.Win.Spread.FpSpread
        Get
            Return ppvwCancelWait
        End Get
        Set(ByVal value As FarPoint.Win.Spread.FpSpread)
            ppvwCancelWait = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【キャンセル済一覧】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppvwCancelDone</returns>
    ''' <remarks><para>作成情報：2015/08/07 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropvwCancelDone() As FarPoint.Win.Spread.FpSpread
        Get
            Return ppvwCancelDone
        End Get
        Set(ByVal value As FarPoint.Win.Spread.FpSpread)
            ppvwCancelDone = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【祝祭日データ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtShukusaijitsu</returns>
    ''' <remarks><para>作成情報：2015/08/11 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtShukusaijitsu() As DataTable
        Get
            Return ppDtShukusaijitsu
        End Get
        Set(ByVal value As DataTable)
            ppDtShukusaijitsu = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【予約一覧データ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtYoyakuList</returns>
    ''' <remarks><para>作成情報：2015/08/10 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtYoyakuList() As DataTable
        Get
            Return ppDtYoyakuList
        End Get
        Set(ByVal value As DataTable)
            ppDtYoyakuList = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【予約未確定一覧データ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtMikakuYoyakuList</returns>
    ''' <remarks><para>作成情報：2015/08/10 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtMikakuYoyakuList() As DataTable
        Get
            Return ppDtMikakuYoyakuList
        End Get
        Set(ByVal value As DataTable)
            ppDtMikakuYoyakuList = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【キャンセル待ちデータ（日付確定のみ）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtDateKakuOnlyCancelWaitList</returns>
    ''' <remarks><para>作成情報：2015/08/10 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtDateKakuOnlyCancelWaitList() As DataTable
        Get
            Return ppDtDateKakuOnlyCancelWaitList
        End Get
        Set(ByVal value As DataTable)
            ppDtDateKakuOnlyCancelWaitList = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【キャンセル待ちデータ（全て）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtNoDateCancelWaitList</returns>
    ''' <remarks><para>作成情報：2015/08/10 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtAllCancelWaitList() As DataTable
        Get
            Return ppDtAllCancelWaitList
        End Get
        Set(ByVal value As DataTable)
            ppDtAllCancelWaitList = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【キャンセル待ち未確定データ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtMikakuCancelWaitList</returns>
    ''' <remarks><para>作成情報：2015/08/10 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtMikakuCancelWaitList() As DataTable
        Get
            Return ppDtMikakuCancelWaitList
        End Get
        Set(ByVal value As DataTable)
            ppDtMikakuCancelWaitList = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【キャンセル済データ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtCancelDoneList</returns>
    ''' <remarks><para>作成情報：2015/08/10 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtCancelDoneList() As DataTable
        Get
            Return ppDtCancelDoneList
        End Get
        Set(ByVal value As DataTable)
            ppDtCancelDoneList = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【その他利用予約一覧データ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtEtcUseYoyakuList</returns>
    ''' <remarks><para>作成情報：2015/08/10 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtEtcUseYoyakuList() As DataTable
        Get
            Return ppDtEtcUseYoyakuList
        End Get
        Set(ByVal value As DataTable)
            ppDtEtcUseYoyakuList = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【休館メンテ一覧データ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtKyuykanMaiteList</returns>
    ''' <remarks><para>作成情報：2015/08/10 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtKyuykanMaiteList() As DataTable
        Get
            Return ppDtKyuykanMaiteList
        End Get
        Set(ByVal value As DataTable)
            ppDtKyuykanMaiteList = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【予約情報（カレンダーからのクリック対応に基づく紐付け用）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDicYoyakuInfo</returns>
    ''' <remarks><para>作成情報：2015/08/20 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDicYoyakuInfo() As Dictionary(Of String, String)
        Get
            Return ppDicYoyakuInfo
        End Get
        Set(ByVal value As Dictionary(Of String, String))
            ppDicYoyakuInfo = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【キャンセル情報（カレンダーからのクリック対応に基づく紐付け用）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDicCancelInfo</returns>
    ''' <remarks><para>作成情報：2015/08/20 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDicCancelInfo() As Dictionary(Of String, String)
        Get
            Return ppDicCancelInfo
        End Get
        Set(ByVal value As Dictionary(Of String, String))
            ppDicCancelInfo = value
        End Set
    End Property

End Class
