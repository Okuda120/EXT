Public Class DataEXTZ0201

    'パラメータ変数宣言
    Private ppStrHolmentDt As String           '対象日
    Private ppStrShisetuKbn As String          '施設区分
    Private ppStrStudioKbn As String           'スタジオ区分
    Private ppStrMnaiyo As String              'メンテナンス内容
    Private ppStrHolmentKbn As String          'メンテナンス区分
    Private ppIntDataCnt As Integer            'データ件数     
    Public PropStrDate As String
    ''' <summary>
    ''' プロパティセット【対象日】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrHolmentDt</returns>
    ''' <remarks><para>作成情報：2015/08/21 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrHolmentDt() As String
        Get
            Return ppStrHolmentDt
        End Get
        Set(ByVal value As String)
            ppStrHolmentDt = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【施設区分】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrShisetuKbn</returns>
    ''' <remarks><para>作成情報：2015/08/21 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrShisetuKbn() As String
        Get
            Return ppStrShisetuKbn
        End Get
        Set(ByVal value As String)
            ppStrShisetuKbn = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【スタジオ区分】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrStudioKbn</returns>
    ''' <remarks><para>作成情報：2015/08/28 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrStudioKbn() As String
        Get
            Return ppStrStudioKbn
        End Get
        Set(ByVal value As String)
            ppStrStudioKbn = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【メンテナンス内容】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrMnaiyo</returns>
    ''' <remarks><para>作成情報：2015/08/28 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrMnaiyo() As String
        Get
            Return ppStrMnaiyo
        End Get
        Set(ByVal value As String)
            ppStrMnaiyo = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【メンテナンス区分】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrHolmentKbn</returns>
    ''' <remarks><para>作成情報：2015/08/28 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropHolmentKbn() As String
        Get
            Return ppStrHolmentKbn
        End Get
        Set(ByVal value As String)
            ppStrHolmentKbn = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【データ件数】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppIntDataCnt</returns>
    ''' <remarks><para>作成情報：2015/08/28 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropIntDataCnt() As Integer
        Get
            Return ppIntDataCnt
        End Get
        Set(ByVal value As Integer)
            ppIntDataCnt = value
        End Set
    End Property

End Class
