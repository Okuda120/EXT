Public Class DataEXTZ0202

    Private ppTxtRiyobiFrom As TextBox              '利用日From
    Private ppTxtRiyobiTo As TextBox                '利用日To
    Private ppStrShisetsuKbn As String              '施設区分
    Private ppStrStudioKbn As String                'スタジオ区分
    Private ppStrTourokuKbn As String               '登録区分
    Private ppStrRegistFlg As String                '登録済フラグ

    Dim ppListRiyobi As New ArrayList               '利用日リスト

    ''' <summary>
    ''' プロパティセット【利用日From】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppTxtRiyobiFrom</returns>
    ''' <remarks><para>作成情報：2015/08/04 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropTxtRiyobiFrom()
        Get
            Return ppTxtRiyobiFrom
        End Get
        Set(ByVal value)
            ppTxtRiyobiFrom = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用日To】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppTxtRiyobiTo</returns>
    ''' <remarks><para>作成情報：2015/08/04 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropTxtRiyobiTo()
        Get
            Return ppTxtRiyobiTo
        End Get
        Set(ByVal value)
            ppTxtRiyobiTo = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【施設区分】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrShisetsuKbn</returns>
    ''' <remarks><para>作成情報：2015/08/04 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrShisetsuKbn()
        Get
            Return ppStrShisetsuKbn
        End Get
        Set(ByVal value)
            ppStrShisetsuKbn = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【スタジオ区分】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrStudioKbn</returns>
    ''' <remarks><para>作成情報：2015/08/04 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrStudioKbn()
        Get
            Return ppStrStudioKbn
        End Get
        Set(ByVal value)
            ppStrStudioKbn = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【登録区分】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrTourokuKbn</returns>
    ''' <remarks><para>作成情報：2015/08/04 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrTourokuKbn()
        Get
            Return ppStrTourokuKbn
        End Get
        Set(ByVal value)
            ppStrTourokuKbn = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【登録済フラグ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRegistFlg</returns>
    ''' <remarks><para>作成情報：2015/08/04 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrRegistFlg()
        Get
            Return ppStrRegistFlg
        End Get
        Set(ByVal value)
            ppStrRegistFlg = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用日リスト】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppListRiyobi</returns>
    ''' <remarks><para>作成情報：2015/08/04 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropListRiyobi()
        Get
            Return ppListRiyobi
        End Get
        Set(ByVal value)
            ppListRiyobi = value
        End Set
    End Property

End Class
