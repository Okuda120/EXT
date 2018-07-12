Public Class CommonDataRiyobi

    Private ppIntSeq As Integer
    Private ppStrShisetuKbn As String
    Private ppStrStudioKbn As String
    Private ppStrYoyakuDt As String
    Private ppStrYoyakuDtDisp As String
    Private ppStrStartTime As String
    Private ppStrEndTime As String
    Private ppStrYoyakuNo As String
    Private ppStrRiyoKeitai As String
    Private ppStrMiteiFlg As String
    Private ppIntTanka As Integer
    Private ppDblBairitu As Double
    Private ppIntSu As Integer
    'Private ppIntRiyoKin As Integer                 ' 2015.12.21 UPD h.hagiwara 
    Private ppIntRiyoKin As Long                     ' 2015.12.21 UPD h.hagiwara 
    Private ppStrRegistFlg As String                '登録済フラグ 1:予約日程表に登録済、0:未登録or予約制御に登録
    Private ppBlnSelect As String                   '選択状態 正式のみで使用

    ''' <summary>
    ''' プロパティセット【SEQ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppIntSeq</returns>
    ''' <remarks><para>作成情報：2015/08/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropIntSeq()
        Get
            Return ppIntSeq
        End Get
        Set(ByVal value)
            ppIntSeq = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【施設区分】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrShisetuKbn</returns>
    ''' <remarks><para>作成情報：2015/08/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrShisetuKbn()
        Get
            Return ppStrShisetuKbn
        End Get
        Set(ByVal value)
            ppStrShisetuKbn = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【スタジオ区分】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrStudioKbn</returns>
    ''' <remarks><para>作成情報：2015/08/07 k.machida
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
    ''' プロパティセット【予約日】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrYoyakuDt</returns>
    ''' <remarks><para>作成情報：2015/08/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrYoyakuDt()
        Get
            Return ppStrYoyakuDt
        End Get
        Set(ByVal value)
            ppStrYoyakuDt = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【予約日(表示用)】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrYoyakuDt</returns>
    ''' <remarks><para>作成情報：2015/08/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrYoyakuDtDisp()
        Get
            Return ppStrYoyakuDtDisp
        End Get
        Set(ByVal value)
            ppStrYoyakuDtDisp = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【開始時間】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrStartTime</returns>
    ''' <remarks><para>作成情報：2015/08/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrStartTime()
        Get
            Return ppStrStartTime
        End Get
        Set(ByVal value)
            ppStrStartTime = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【終了時間】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrEndTime</returns>
    ''' <remarks><para>作成情報：2015/08/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrEndTime()
        Get
            Return ppStrEndTime
        End Get
        Set(ByVal value)
            ppStrEndTime = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【予約番号】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrYoyakuNo</returns>
    ''' <remarks><para>作成情報：2015/08/07 k.machida
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
    ''' プロパティセット【利用形態】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyoKeitai</returns>
    ''' <remarks><para>作成情報：2015/08/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrRiyoKeitai()
        Get
            Return ppStrRiyoKeitai
        End Get
        Set(ByVal value)
            ppStrRiyoKeitai = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【未定フラグ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrMiteiFlg</returns>
    ''' <remarks><para>作成情報：2015/08/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrMiteiFlg()
        Get
            Return ppStrMiteiFlg
        End Get
        Set(ByVal value)
            ppStrMiteiFlg = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【単価】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppIntTanka</returns>
    ''' <remarks><para>作成情報：2015/08/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropIntTanka()
        Get
            Return ppIntTanka
        End Get
        Set(ByVal value)
            ppIntTanka = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【倍率】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDblBairitu</returns>
    ''' <remarks><para>作成情報：2015/08/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDblBairitu()
        Get
            Return ppDblBairitu
        End Get
        Set(ByVal value)
            ppDblBairitu = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【数】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppIntSu</returns>
    ''' <remarks><para>作成情報：2015/08/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropIntSu()
        Get
            Return ppIntSu
        End Get
        Set(ByVal value)
            ppIntSu = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用金額】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppIntRiyoKin</returns>
    ''' <remarks><para>作成情報：2015/08/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropIntRiyoKin()
        Get
            Return ppIntRiyoKin
        End Get
        Set(ByVal value)
            ppIntRiyoKin = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【登録済フラグ】
    ''' </summary>
    ''' <value>DB登録済データか判定する。</value>
    ''' <returns>ppStrRegistFlg</returns>
    ''' <remarks><para>作成情報：2015/08/07 k.machida
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
    ''' プロパティセット【選択状態】
    ''' </summary>
    ''' <value>選択状態、正式予約のみでしようする</value>
    ''' <returns>ppBlnSelect</returns>
    ''' <remarks><para>作成情報：2015/08/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropBlnSelect()
        Get
            Return ppBlnSelect
        End Get
        Set(ByVal value)
            ppBlnSelect = value
        End Set
    End Property

End Class
