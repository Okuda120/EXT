Public Class CommonDataCancel

    Private ppIntSeq As Integer
    Private ppStrShisetuKbn As String
    Private ppStrStudioKbn As String
    Private ppStrKibobiKbn As String
    Private ppStrCancelYm As String
    Private ppStrCancelDt As String
    Private ppStrCancelDtDisp As String
    Private ppStrRiyoKeitai As String
    Private ppStrMiteiFlg As String
    Private ppStrStartTime As String
    Private ppStrEndTime As String
    Private ppStrMemo As String
    Private ppStrCancelWaitNo As String
    Private ppIntWakuNo As Integer
    Private ppStrRegistFlg As String                '登録済フラグ 1:予約日程表に登録済、0:未登録or予約制御に登録

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
    ''' プロパティセット【未定フラグ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrDateUmu</returns>
    ''' <remarks><para>作成情報：2015/08/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrKibobiKbn()
        Get
            Return ppStrKibobiKbn
        End Get
        Set(ByVal value)
            ppStrKibobiKbn = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【予約日】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrCancelYm</returns>
    ''' <remarks><para>作成情報：2015/08/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrCancelYm()
        Get
            Return ppStrCancelYm
        End Get
        Set(ByVal value)
            ppStrCancelYm = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【予約日】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrCancelDt</returns>
    ''' <remarks><para>作成情報：2015/08/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrCancelDt()
        Get
            Return ppStrCancelDt
        End Get
        Set(ByVal value)
            ppStrCancelDt = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【予約日(表示用)】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrCancelDtDisp</returns>
    ''' <remarks><para>作成情報：2015/08/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrCancelDtDisp()
        Get
            Return ppStrCancelDtDisp
        End Get
        Set(ByVal value)
            ppStrCancelDtDisp = value
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
    ''' プロパティセット【Memo】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrMemo</returns>
    ''' <remarks><para>作成情報：2015/08/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrMemo()
        Get
            Return ppStrMemo
        End Get
        Set(ByVal value)
            ppStrMemo = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【予約番号】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrCancelWaitNo</returns>
    ''' <remarks><para>作成情報：2015/08/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrCancelWaitNo()
        Get
            Return ppStrCancelWaitNo
        End Get
        Set(ByVal value)
            ppStrCancelWaitNo = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【枠番】
    ''' </summary>
    ''' <value>DB登録済データか判定する。</value>
    ''' <returns>ppIntWakuNo</returns>
    ''' <remarks><para>作成情報：2015/08/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropIntWakuNo()
        Get
            Return ppIntWakuNo
        End Get
        Set(ByVal value)
            ppIntWakuNo = value
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

End Class
