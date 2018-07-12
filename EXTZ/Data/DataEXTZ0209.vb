Imports FarPoint.Win.Spread

Public Class DataEXTZ0209

    Private ppStrYoyakuNo As String
    Private ppStrSaijiNm As String
    Private ppStrShisetu As String
    Private ppStrRiyobi As String
    Private ppStrRiyobiDisp As String
    'Private ppIntTotal As Integer                        ' 2015.12.21 UPD h.hagiwara
    Private ppIntTotal As Long                            ' 2015.12.21 UPD h.hagiwara
    Private ppStrFutaiTotalNm As String
    Private ppFutaiDetailTable As DataTable '
    Private ppFutaiBunruiTable As DataTable '
    Private ppFutaiMstTable As DataTable '
    'Private ppIntTotalTax As Integer                 ' 2015.11.13 ADD h.hagiwara
    Private ppIntTotalTax As Long                     ' 2015.11.21 UPD h.hagiwara
    Private ppLngTotal As Long                        ' 2015.12.21 ADD h.hagiwara

    'シート
    Private ppVwFutaiSelecSheet As FpSpread          ' 2015.12.01 ADD h.hagiwara

    ''' <summary>
    ''' プロパティセット【予約No】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrYoyakuNo</returns>
    ''' <remarks><para>作成情報：2015/08/10 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrYoyakuNo()
        Get
            Return ppStrYoyakuNo
        End Get
        Set(ByVal value)
            ppStrYoyakuNo = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【催事名】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrSaijiNm</returns>ppStrStudioKbn
    Public Property PropStrSaijiNm()
        Get
            Return ppStrSaijiNm
        End Get
        Set(ByVal value)
            ppStrSaijiNm = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【施設区分】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrShisetu</returns>
    ''' <remarks><para>作成情報：2015/08/30 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrShisetu()
        Get
            Return ppStrShisetu
        End Get
        Set(ByVal value)
            ppStrShisetu = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用日】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyobi</returns>
    ''' <remarks><para>作成情報：2015/08/30 k.machida
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
    ''' プロパティセット【利用日】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyobiDisp</returns>
    ''' <remarks><para>作成情報：2015/08/30 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrRiyobiDisp()
        Get
            Return ppStrRiyobiDisp
        End Get
        Set(ByVal value)
            ppStrRiyobiDisp = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【Total】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppIntTotal</returns>
    ''' <remarks><para>作成情報：2015/08/30 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropIntTotal()
        Get
            Return ppIntTotal
        End Get
        Set(ByVal value)
            ppIntTotal = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【付帯設備名称(結合)】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrFutaiTotalNm</returns>
    ''' <remarks><para>作成情報：2015/08/31 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrFutaiTotalNm()
        Get
            Return ppStrFutaiTotalNm
        End Get
        Set(ByVal value)
            ppStrFutaiTotalNm = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【付帯設備明細】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppFutaiDetailTable</returns>
    ''' <remarks><para>作成情報：2015/08/30 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropFutaiDetailTable()
        Get
            Return ppFutaiDetailTable
        End Get
        Set(ByVal value)
            ppFutaiDetailTable = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【付帯設備分類】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppFutaiBunruiTable</returns>
    ''' <remarks><para>作成情報：2015/08/30 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropFutaiBunruiTable()
        Get
            Return ppFutaiBunruiTable
        End Get
        Set(ByVal value)
            ppFutaiBunruiTable = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【付帯設備マスタ情報】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppFutaiMstTable</returns>
    ''' <remarks><para>作成情報：2015/08/30 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropFutaiMstTable()
        Get
            Return ppFutaiMstTable
        End Get
        Set(ByVal value)
            ppFutaiMstTable = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【Total】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppIntTotalTax</returns>
    ''' <remarks><para>作成情報：2015.11.13 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropIntTotalTax()
        Get
            Return ppIntTotalTax
        End Get
        Set(ByVal value)
            ppIntTotalTax = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【ppVwGroupingSheet】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppVwFutaiSelecSheet</returns>
    ''' <remarks><para>作成情報：2015.12.01 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropVwFutaiSelecSheet() As FpSpread
        Get
            Return ppVwFutaiSelecSheet
        End Get
        Set(ByVal value As FpSpread)
            ppVwFutaiSelecSheet = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【予約番号の付帯設備合計金額】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppLngTotal</returns>
    ''' <remarks><para>作成情報：2015.12.21 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropLngTotal()
        Get
            Return ppLngTotal
        End Get
        Set(ByVal value)
            ppLngTotal = value
        End Set
    End Property

End Class
