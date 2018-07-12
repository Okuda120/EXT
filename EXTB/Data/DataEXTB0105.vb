Public Class DataEXTB0105

    'パラメータ変数宣言
    Private ppStrRiyoDt As String      '利用日
    Private ppStrStartTime As String   '開始時間
    Private ppStrEndTime As String     '終了時間
    Private ppStrRiyoBasho As String   '場所
    Private ppStrRiyoYoto As String    '用途
    Private ppStrRiyosha As String     '利用者
    Private ppIntDataCnt As Integer    'データ件数     


    ''' <summary>
    ''' プロパティセット【利用日】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyoDt</returns>
    ''' <remarks><para>作成情報：2015/08/19 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrRiyoDt() As String
        Get
            Return ppStrRiyoDt
        End Get
        Set(ByVal value As String)
            ppStrRiyoDt = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【開始時間】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrStartTime</returns>
    ''' <remarks><para>作成情報：2015/08/27 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrStartTime() As String
        Get
            Return ppStrStartTime
        End Get
        Set(ByVal value As String)
            ppStrStartTime = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【終了時間】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrEndTime</returns>
    ''' <remarks><para>作成情報：2015/08/27 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrEndTime() As String
        Get
            Return ppStrEndTime
        End Get
        Set(ByVal value As String)
            ppStrEndTime = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【場所】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyoBasho</returns>
    ''' <remarks><para>作成情報：2015/08/27 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrRiyoBasho() As String
        Get
            Return ppStrRiyoBasho
        End Get
        Set(ByVal value As String)
            ppStrRiyoBasho = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【用途】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyoYoto</returns>
    ''' <remarks><para>作成情報：2015/08/27 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrRiyoYoto() As String
        Get
            Return ppStrRiyoYoto
        End Get
        Set(ByVal value As String)
            ppStrRiyoYoto = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyosha</returns>
    ''' <remarks><para>作成情報：2015/08/27 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrRiyosha() As String
        Get
            Return ppStrRiyosha
        End Get
        Set(ByVal value As String)
            ppStrRiyosha = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【データ件数】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppIntDataCnt</returns>
    ''' <remarks><para>作成情報：2015/08/27 h.endo
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
