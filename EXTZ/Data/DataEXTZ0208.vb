Imports FarPoint.Win.Spread

Public Class DataEXTZ0208

    Private ppStrAiteCd As String
    Private ppStrAiteNm As String
    Private ppStrNyukinYoteiFrom As String
    Private ppStrNyukinYoteiTo As String
    Private ppStrNyukinFrom As String
    Private ppStrNyukinTo As String
    Private ppStrSeikyuNo As String
    Private ppStrSeikyuIraiNo As String
    Private ppDtResult As DataTable
    Private ppStrResAiteCd As String
    Private ppStrResAiteNm As String
    Private ppStrResNyukinYoteiDt As String
    Private ppStrResNyukinDt As String
    Private ppIntResNyukinKin As Integer
    Private ppStrResSeikyuDt As String
    Private ppStrResInputDt As String
    Private ppStrResSeikyuNo As String
    Private ppStrResSeikyuIraiNo As String
    Private ppBlnChangeFlg As Boolean
    Private ppIntResNyukinLink As Integer
    Private ppStrNyukinInputFlg As String
    Private ppDrBillReq As DataRow        '請求
    Private ppIntCheckRow As Integer      'チェック行番号
    Private fbResult As FpSpread          'スプレッドシート

    ''' <summary>
    ''' プロパティセット【検索キー：相手CD】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrAiteCd</returns>
    ''' <remarks><para>作成情報：2015/09/02 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrAiteCd()
        Get
            Return ppStrAiteCd
        End Get
        Set(ByVal value)
            ppStrAiteCd = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【検索キー：相手名称】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrAiteNm</returns>
    ''' <remarks><para>作成情報：2015/09/02 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrAiteNm()
        Get
            Return ppStrAiteNm
        End Get
        Set(ByVal value)
            ppStrAiteNm = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【検索キー：入金予定日FROM】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrNyukinYoteiFrom</returns>
    ''' <remarks><para>作成情報：2015/09/02 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrNyukinYoteiFrom()
        Get
            Return ppStrNyukinYoteiFrom
        End Get
        Set(ByVal value)
            ppStrNyukinYoteiFrom = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【検索キー：入金予定日TO】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrNyukinYoteiTo</returns>
    ''' <remarks><para>作成情報：2015/09/02 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrNyukinYoteiTo()
        Get
            Return ppStrNyukinYoteiTo
        End Get
        Set(ByVal value)
            ppStrNyukinYoteiTo = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【検索キー：入金日FROM】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrNyukinFrom</returns>
    ''' <remarks><para>作成情報：2015/09/02 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrNyukinFrom()
        Get
            Return ppStrNyukinFrom
        End Get
        Set(ByVal value)
            ppStrNyukinFrom = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【検索キー：入金日TO】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrNyukinTo</returns>
    ''' <remarks><para>作成情報：2015/09/02 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrNyukinTo()
        Get
            Return ppStrNyukinTo
        End Get
        Set(ByVal value)
            ppStrNyukinTo = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【検索キー：請求NO】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrSeikyuNo</returns>
    ''' <remarks><para>作成情報：2015/09/02 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrSeikyuNo()
        Get
            Return ppStrSeikyuNo
        End Get
        Set(ByVal value)
            ppStrSeikyuNo = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【検索キー：請求依頼番号】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrSeikyuIraiNo</returns>
    ''' <remarks><para>作成情報：2015/09/02 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrSeikyuIraiNo()
        Get
            Return ppStrSeikyuIraiNo
        End Get
        Set(ByVal value)
            ppStrSeikyuIraiNo = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【検索結果】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtResult</returns>
    ''' <remarks><para>作成情報：2015/09/01 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtResult()
        Get
            Return ppDtResult
        End Get
        Set(ByVal value)
            ppDtResult = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【選択行：相手CD】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrResAiteCd</returns>
    ''' <remarks><para>作成情報：2015/09/02 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrResAiteCd()
        Get
            Return ppStrResAiteCd
        End Get
        Set(ByVal value)
            ppStrResAiteCd = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【選択行：相手名称】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrResAiteNm</returns>
    ''' <remarks><para>作成情報：2015/09/02 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrResAiteNm()
        Get
            Return ppStrResAiteNm
        End Get
        Set(ByVal value)
            ppStrResAiteNm = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【選択行：入金予定日】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrResNyukinYoteiDt</returns>
    ''' <remarks><para>作成情報：2015/09/02 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrResNyukinYoteiDt()
        Get
            Return ppStrResNyukinYoteiDt
        End Get
        Set(ByVal value)
            ppStrResNyukinYoteiDt = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【選択行：入金日】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrResNyukinDt</returns>
    ''' <remarks><para>作成情報：2015/09/02 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrResNyukinDt()
        Get
            Return ppStrResNyukinDt
        End Get
        Set(ByVal value)
            ppStrResNyukinDt = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【選択行：入金額(請求金)】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppIntResNyukinKin</returns>
    ''' <remarks><para>作成情報：2015/09/02 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropIntResNyukinKin()
        Get
            Return ppIntResNyukinKin
        End Get
        Set(ByVal value)
            ppIntResNyukinKin = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【選択行：請求日】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrResSeikyuDt</returns>
    ''' <remarks><para>作成情報：2015/09/02 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrResSeikyuDt()
        Get
            Return ppStrResSeikyuDt
        End Get
        Set(ByVal value)
            ppStrResSeikyuDt = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【選択行：入力日】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrResInputDt</returns>
    ''' <remarks><para>作成情報：2015/09/02 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrResInputDt()
        Get
            Return ppStrResInputDt
        End Get
        Set(ByVal value)
            ppStrResInputDt = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【選択行：請求NO】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrResSekikyuNo</returns>
    ''' <remarks><para>作成情報：2015/09/02 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrResSeikyuNo()
        Get
            Return ppStrResSeikyuNo
        End Get
        Set(ByVal value)
            ppStrResSeikyuNo = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【選択行：請求依頼番号】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrResSeikyuIraiNo</returns>
    ''' <remarks><para>作成情報：2015/09/02 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrResSeikyuIraiNo()
        Get
            Return ppStrResSeikyuIraiNo
        End Get
        Set(ByVal value)
            ppStrResSeikyuIraiNo = value
        End Set
    End Property


    ''' <summary>
    ''' プロパティセット【変更フラグ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppBlnChangeFlg</returns>
    ''' <remarks><para>作成情報：2015/09/03 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropBlnChangeFlg()
        Get
            Return ppBlnChangeFlg
        End Get
        Set(ByVal value)
            ppBlnChangeFlg = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【選択行：入金LINKNO】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppIntResNyukinLink</returns>
    ''' <remarks><para>作成情報：2015/09/02 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropIntResNyukinLink()
        Get
            Return ppIntResNyukinLink
        End Get
        Set(ByVal value)
            ppIntResNyukinLink = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【入金入力済フラグ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrNyukinInputFlg</returns>
    ''' <remarks><para>作成情報：2015/08/27 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property propStrNyukinInputFlg()
        Get
            Return ppStrNyukinInputFlg
        End Get
        Set(ByVal value)
            ppStrNyukinInputFlg = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【データ行:請求】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDrBillReq</returns>
    ''' <remarks><para>作成情報：2015/09/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDrBillReq()
        Get
            Return ppDrBillReq
        End Get
        Set(ByVal value)
            ppDrBillReq = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【スプレッドシート】
    ''' </summary>
    ''' <value></value>
    ''' <returns>fbResult</returns>
    ''' <remarks><para>作成情報：2015/11/30 hayabuchi
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropResult() As FpSpread
        Get
            Return fbResult
        End Get
        Set(ByVal value As FpSpread)
            fbResult = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【チェック行番号】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppIntCheckRow</returns>
    ''' <remarks><para>作成情報：2015/11/30 hayabuchi
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropIntCheckRow() As Integer
        Get
            Return ppIntCheckRow
        End Get
        Set(ByVal value As Integer)
            ppIntCheckRow = value
        End Set
    End Property
End Class
