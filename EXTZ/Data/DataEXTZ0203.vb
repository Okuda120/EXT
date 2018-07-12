Public Class DataEXTZ0203

    Private ppStrYoyakuNo As String
    Private ppStrSeq As String
    Private ppStrSeikyuIraiNo As String
    Private ppStrSeikyuDt As String
    Private ppStrNyukinYoteiDt As String
    Private ppStrKakuteiKin As String
    Private ppStrChoseiKin As String
    Private ppStrShokei As String
    Private ppStrTaxKin As String
    Private ppStrSeikyuKin As String
    Private ppStrSeikyuNaiyo As String
    Private ppStrSeikyuNaiyoNm As String
    Private ppStrAiteCd As String
    Private ppStrAiteNm As String
    Private ppStrNyukinKbn As String
    Private ppStrSeikyuTitle1 As String
    Private ppStrSeikyuTitle2 As String
    Private ppStrNyukinDt As String
    Private ppStrNyukinKin As String
    Private ppStrSeikyuInputFlg As String
    Private ppStrSeikyuIraiFlg As String
    Private ppStrNyukinInputFlg As String
    Private ppStrNyukinLinkNo As String
    Private ppStrGetSeikyuInputFlg As String                         ' 2016.02.03 ADD h.hagiwara

    Private ppDblTax As Double      'TAX
    Private ppIntEditLine As String '選択行番号
    Private ppDsBillReq As DataSet  '請求/入金/EXAS入金
    Private ppDtExasRiyoryo As DataTable  'EXASPRJ利用料
    Private ppDtExasFutai As DataTable  'EXASPRJ付帯

    ''' <summary>
    ''' プロパティセット【予約NO】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrYoyakuNo</returns>
    ''' <remarks><para>作成情報：2015/08/27 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property propStrYoyakuNo()
        Get
            Return ppStrYoyakuNo
        End Get
        Set(ByVal value)
            ppStrYoyakuNo = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【SEQ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrSeq</returns>
    ''' <remarks><para>作成情報：2015/08/27 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property propStrSeq()
        Get
            Return ppStrSeq
        End Get
        Set(ByVal value)
            ppStrSeq = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【請求依頼番号】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrSeikyuIraiNo</returns>
    ''' <remarks><para>作成情報：2015/08/27 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property propStrSeikyuIraiNo()
        Get
            Return ppStrSeikyuIraiNo
        End Get
        Set(ByVal value)
            ppStrSeikyuIraiNo = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【請求日】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrSeikyuDt</returns>
    ''' <remarks><para>作成情報：2015/08/27 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property propStrSeikyuDt()
        Get
            Return ppStrSeikyuDt
        End Get
        Set(ByVal value)
            ppStrSeikyuDt = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【入金予定日】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrNyukinYoteiDt</returns>
    ''' <remarks><para>作成情報：2015/08/27 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property propStrNyukinYoteiDt()
        Get
            Return ppStrNyukinYoteiDt
        End Get
        Set(ByVal value)
            ppStrNyukinYoteiDt = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【確定額】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrKakuteiKin</returns>
    ''' <remarks><para>作成情報：2015/08/27 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property propStrKakuteiKin()
        Get
            Return ppStrKakuteiKin
        End Get
        Set(ByVal value)
            ppStrKakuteiKin = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【調整額】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrChoseiKin</returns>
    ''' <remarks><para>作成情報：2015/08/27 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property propStrChoseiKin()
        Get
            Return ppStrChoseiKin
        End Get
        Set(ByVal value)
            ppStrChoseiKin = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【小計】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrShokei</returns>
    ''' <remarks><para>作成情報：2015/08/27 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property propStrShokei()
        Get
            Return ppStrShokei
        End Get
        Set(ByVal value)
            ppStrShokei = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【消費税額】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrTaxKin</returns>
    ''' <remarks><para>作成情報：2015/08/27 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property propStrTaxKin()
        Get
            Return ppStrTaxKin
        End Get
        Set(ByVal value)
            ppStrTaxKin = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【請求金額】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrSeikyuKin</returns>
    ''' <remarks><para>作成情報：2015/08/27 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property propStrSeikyuKin()
        Get
            Return ppStrSeikyuKin
        End Get
        Set(ByVal value)
            ppStrSeikyuKin = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【請求内容コード】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrSeikyuNaiyo</returns>
    ''' <remarks><para>作成情報：2015/08/27 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property propStrSeikyuNaiyo()
        Get
            Return ppStrSeikyuNaiyo
        End Get
        Set(ByVal value)
            ppStrSeikyuNaiyo = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【請求内容】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrSeikyuNaiyoNm</returns>
    ''' <remarks><para>作成情報：2015/08/27 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property propStrSeikyuNaiyoNm()
        Get
            Return ppStrSeikyuNaiyoNm
        End Get
        Set(ByVal value)
            ppStrSeikyuNaiyoNm = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【相手先コード】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrAiteCd</returns>
    ''' <remarks><para>作成情報：2015/08/27 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property propStrAiteCd()
        Get
            Return ppStrAiteCd
        End Get
        Set(ByVal value)
            ppStrAiteCd = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【相手先名】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrAiteNm</returns>
    ''' <remarks><para>作成情報：2015/08/27 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property propStrAiteNm()
        Get
            Return ppStrAiteNm
        End Get
        Set(ByVal value)
            ppStrAiteNm = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【入金種別】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrNyukinKbn</returns>
    ''' <remarks><para>作成情報：2015/08/27 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property propStrNyukinKbn()
        Get
            Return ppStrNyukinKbn
        End Get
        Set(ByVal value)
            ppStrNyukinKbn = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【請求書タイトル１】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrSeikyuTitle1</returns>
    ''' <remarks><para>作成情報：2015/08/27 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property propStrSeikyuTitle1()
        Get
            Return ppStrSeikyuTitle1
        End Get
        Set(ByVal value)
            ppStrSeikyuTitle1 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【請求書タイトル２】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrSeikyuTitle2</returns>
    ''' <remarks><para>作成情報：2015/08/27 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property propStrSeikyuTitle2()
        Get
            Return ppStrSeikyuTitle2
        End Get
        Set(ByVal value)
            ppStrSeikyuTitle2 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【入金日】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrNyukinDt</returns>
    ''' <remarks><para>作成情報：2015/08/27 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property propStrNyukinDt()
        Get
            Return ppStrNyukinDt
        End Get
        Set(ByVal value)
            ppStrNyukinDt = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【入金額】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrNyukinKin</returns>
    ''' <remarks><para>作成情報：2015/08/27 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property propStrNyukinKin()
        Get
            Return ppStrNyukinKin
        End Get
        Set(ByVal value)
            ppStrNyukinKin = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【請求入力済フラグ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrSeikyuInputFlg</returns>
    ''' <remarks><para>作成情報：2015/08/27 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property propStrSeikyuInputFlg()
        Get
            Return ppStrSeikyuInputFlg
        End Get
        Set(ByVal value)
            ppStrSeikyuInputFlg = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【EXAS請求依頼済フラグ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrSeikyuIraiFlg</returns>
    ''' <remarks><para>作成情報：2015/08/27 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property propStrSeikyuIraiFlg()
        Get
            Return ppStrSeikyuIraiFlg
        End Get
        Set(ByVal value)
            ppStrSeikyuIraiFlg = value
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
    ''' プロパティセット【EXAS入金リンク番号】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrNyukinLinkNo</returns>
    ''' <remarks><para>作成情報：2015/08/27 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property propStrNyukinLinkNo()
        Get
            Return ppStrNyukinLinkNo
        End Get
        Set(ByVal value)
            ppStrNyukinLinkNo = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【消費税】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDblTax</returns>
    ''' <remarks><para>作成情報：2015/08/28 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property propDblTax()
        Get
            Return ppDblTax
        End Get
        Set(ByVal value)
            ppDblTax = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【選択行】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppIntEditLine</returns>
    ''' <remarks><para>作成情報：2015/08/28 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property propIntEditLine()
        Get
            Return ppIntEditLine
        End Get
        Set(ByVal value)
            ppIntEditLine = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【データセット:請求】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDsBillReq</returns>
    ''' <remarks><para>作成情報：2015/08/25 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDsBillReq()
        Get
            Return ppDsBillReq
        End Get
        Set(ByVal value)
            ppDsBillReq = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【データテーブル:EXAS利用料】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtExasRiyoryo</returns>
    ''' <remarks><para>作成情報：2015/08/25 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtExasRiyoryo()
        Get
            Return ppDtExasRiyoryo
        End Get
        Set(ByVal value)
            ppDtExasRiyoryo = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【データテーブル:EXAS付帯】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtExasFutai</returns>
    ''' <remarks><para>作成情報：2015/08/25 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtExasFutai()
        Get
            Return ppDtExasFutai
        End Get
        Set(ByVal value)
            ppDtExasFutai = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【ＤＢ請求入力済フラグ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrGetSeikyuInputFlg</returns>
    ''' <remarks><para>作成情報：2016.02.03 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property propStrGetSeikyuInputFlg()
        Get
            Return ppStrGetSeikyuInputFlg
        End Get
        Set(ByVal value)
            ppStrGetSeikyuInputFlg = value
        End Set
    End Property

End Class
