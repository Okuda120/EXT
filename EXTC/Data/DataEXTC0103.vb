Imports FarPoint.Win.Spread

Public Class DataEXTC0103

    '予約番号
    Private ppStrReserveNo As String
    Private ppLblYoyakuNo As Label

    '出力者
    Private ppLoginUser As String

    Private ppDtReportData As DataTable                 '帳票出力用データテーブル
    Private ppVwStudioSheet As FpSpread                 '出力用シート
    Private ppDtUseDetails_Output As DataTable          '利用明細
    Private ppDtUseDetailsNoTax_Output As DataTable     '利用明細(税込分)
    Private ppBlnGeneralFlg As Boolean
    Private ppBlnSumFlg As Boolean                      '合算フラグ
    Private ppStrClickedIncident As String              '付帯設備スプレッドシートにてボタンがクリックされた行の日付
    Private ppStrClickedRiyobi As String                ' 利用日時でクリックされた行の日付         ' 2015.12.16 ADD h.hagiwara
    Private ppBlnNoTaxFlg As Boolean                    '税無フラグ
    Private ppIntTax As Integer                         '消費税
    '請求データＣＳＶ作成用
    Private ppStrCalculateDay_Output As String          '税率計算日
    Private ppStrGrpKey As String                       ' 集計キー
    Private ppStrNoTaxFlg As String                     ' 税無フラグ
    Private ppStrAitecd As String                       ' 相手先コード
    Private ppStrPostno As String                       ' 郵便番号
    Private ppStrAddr1 As String                        ' 住所１
    Private ppStrAddr2 As String                        ' 住所２
    Private ppStrAitenm As String                       ' 相手先名


    ''' <summary>
    ''' プロパティセット【帳票：承認番号(予約番号)】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrReserveNo</returns>
    ''' <remarks><para>作成情報：2015/08/10 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrReserveNo() As String
        Get
            Return ppStrReserveNo
        End Get
        Set(ByVal value As String)
            ppStrReserveNo = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【ppDtReportData】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtReportData</returns>
    ''' <remarks><para>作成情報：2015/08/31 yu.satoh
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtReportData() As DataTable
        Get
            Return ppDtReportData
        End Get
        Set(ByVal value As DataTable)
            ppDtReportData = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【帳票：利用明細データテーブル】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtUseDetails_Output</returns>
    ''' <remarks><para>作成情報：2015/08/14 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtUseDetails_Output() As DataTable
        Get
            Return ppDtUseDetails_Output
        End Get
        Set(ByVal value As DataTable)
            ppDtUseDetails_Output = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【帳票：利用明細（税込分）データテーブル】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtUseDetailsNoTax_Output</returns>
    ''' <remarks><para>作成情報：2015/08/14 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtUseDetailsNoTax_Output() As DataTable
        Get
            Return ppDtUseDetailsNoTax_Output
        End Get
        Set(ByVal value As DataTable)
            ppDtUseDetailsNoTax_Output = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【合算フラグ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppBlnSumFlg</returns>
    ''' <remarks><para>作成情報：2015/08/17 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropBlnSumFlg() As Boolean
        Get
            Return ppBlnSumFlg
        End Get
        Set(ByVal value As Boolean)
            ppBlnSumFlg = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【付帯設備スプレッドシートにてボタンがクリックされた行の日付】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrClickedIncident</returns>
    ''' <remarks><para>作成情報：2015/08/18 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrClickedIncident() As String
        Get
            Return ppStrClickedIncident
        End Get
        Set(ByVal value As String)
            ppStrClickedIncident = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【社内社外フラグ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppBlnGeneralFlg</returns>
    ''' <remarks><para>作成情報：2015/09/01 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropBlnGeneralFlg() As Boolean
        Get
            Return ppBlnGeneralFlg
        End Get
        Set(ByVal value As Boolean)
            ppBlnGeneralFlg = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【税無フラグ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppBlnNoTaxFlg</returns>
    ''' <remarks><para>作成情報：2015/08/14 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropBlnNoTaxFlg() As Boolean
        Get
            Return ppBlnNoTaxFlg
        End Get
        Set(ByVal value As Boolean)
            ppBlnNoTaxFlg = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【消費税】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppIntTax</returns>
    ''' <remarks><para>作成情報：2015/09/02 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropIntTax() As Integer
        Get
            Return ppIntTax
        End Get
        Set(ByVal value As Integer)
            ppIntTax = value
        End Set
    End Property



    ''' <summary>
    ''' プロパティセット【ppLblYoyakuNo】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppLblYoyakuNo</returns>
    ''' <remarks><para>作成情報：2015/09/01 yu.satoh
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropLblYoyakuNo() As Label
        Get
            Return ppLblYoyakuNo
        End Get
        Set(ByVal value As Label)
            ppLblYoyakuNo = value
        End Set
    End Property


    ''' <summary>
    ''' プロパティセット【ppVwStudioSheet】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppVwStudioSheet</returns>
    ''' <remarks><para>作成情報：2015/09/01 yu.satoh
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropVwStudioSheet() As FpSpread
        Get
            Return ppVwStudioSheet
        End Get
        Set(ByVal value As FpSpread)
            ppVwStudioSheet = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【ppLoginUser】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppLoginUser</returns>
    ''' <remarks><para>作成情報：2015/09/03 yu.satoh
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropLoginUser() As String
        Get
            Return ppLoginUser
        End Get
        Set(ByVal value As String)
            ppLoginUser = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【帳票：税率計算日】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrCalculateDay_Output</returns>
    ''' <remarks><para>作成情報：2015/08/14 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrCalculateDay_Output() As String
        Get
            Return ppStrCalculateDay_Output
        End Get
        Set(ByVal value As String)
            ppStrCalculateDay_Output = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【EXAS請求情報作成時の消費税率判定用集計キー】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrGrpKey</returns>
    ''' <remarks><para>作成情報：2015/11/11 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrGrpKey() As String
        Get
            Return ppStrGrpKey
        End Get
        Set(ByVal value As String)
            ppStrGrpKey = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【EXAS請求情報作成時の消費税率判定用税抜フラグ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrNoTaxFlg</returns>
    ''' <remarks><para>作成情報：2015/11/11 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrNoTaxFlg() As String
        Get
            Return ppStrNoTaxFlg
        End Get
        Set(ByVal value As String)
            ppStrNoTaxFlg = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【EXAS請求情報用：相手先コード】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrAitecd</returns>
    ''' <remarks><para>作成情報：2015/11/11 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrAitecd() As String
        Get
            Return ppStrAitecd
        End Get
        Set(ByVal value As String)
            ppStrAitecd = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【EXAS請求情報用：郵便番号】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrPostno</returns>
    ''' <remarks><para>作成情報：2015/11/11 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrPostno() As String
        Get
            Return ppStrPostno
        End Get
        Set(ByVal value As String)
            ppStrPostno = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【EXAS請求情報用：住所１】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrAddr1</returns>
    ''' <remarks><para>作成情報：2015/11/11 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrAddr1() As String
        Get
            Return ppStrAddr1
        End Get
        Set(ByVal value As String)
            ppStrAddr1 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【EXAS請求情報用：住所２】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrAddr2</returns>
    ''' <remarks><para>作成情報：2015/11/11 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrAddr2() As String
        Get
            Return ppStrAddr2
        End Get
        Set(ByVal value As String)
            ppStrAddr2 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【EXAS請求情報用：相手先名】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrAitenm</returns>
    ''' <remarks><para>作成情報：2015/11/11 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrAitenm() As String
        Get
            Return ppStrAitenm
        End Get
        Set(ByVal value As String)
            ppStrAitenm = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用日時でクリックされた行の日付】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrClickedRiyobi</returns>
    ''' <remarks><para>作成情報：2015.12.16.h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrClickedRiyobi() As String
        Get
            Return ppStrClickedRiyobi
        End Get
        Set(ByVal value As String)
            ppStrClickedRiyobi = value
        End Set
    End Property

End Class
