Public Class DataEXTZ0101

    'パラメータ変数宣言
    Private ppStrShisetsuKbn As String      '施設区分
    Private ppStrRiyoDtFrom As String       '利用日（From）
    Private ppStrRiyoDtTo As String         '利用日（To）
    Private ppStrRiyoNm As String           '利用者名
    Private ppStrRiyoKana As String         '利用者名カナ
    Private ppStrSaijiNm As String          '催事名
    Private ppStrYoyakuNo As String         '予約NO
    Private ppBlnMikanryo As Boolean        '未完了フラグ
    Private ppvwYoyakuTheatre As FarPoint.Win.Spread.FpSpread       '予約一覧（シアター）
    Private ppvwYoyakuStudio As FarPoint.Win.Spread.FpSpread        '予約一覧（スタジオ）
    Private ppDtYoyaku As DataTable         '予約データ
    Private ppBtnDecision As Button         '選択確定ボタン 
    Private ppFrmYoyaku As Form             '予約一覧フォーム
    Private ppStrSelectMode As String       '選択モード
    Private ppStrRtnYoyakuNo As String      '遷移元戻し予約NO
    Private ppStrRtnSaijiNm As String       '遷移元戻し催事名
    Private ppStrSelectShisetu As String    '選択モード(施設区分)        ' 2015.12.15 ADD h.hagiwara
    'データ
    Private ppIntCheckRow As Integer        'チェックされた行番号　　2015.12.01 y.ozawa 単一チェック対応

    ''' <summary>
    ''' プロパティセット【施設区分】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrShisetsuKbn</returns>
    ''' <remarks><para>作成情報：2015/09/10 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrShisetsuKbn() As String
        Get
            Return ppStrShisetsuKbn
        End Get
        Set(ByVal value As String)
            ppStrShisetsuKbn = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用日（From）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyoDtFrom</returns>
    ''' <remarks><para>作成情報：2015/09/09 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrRiyoDtFrom() As String
        Get
            Return ppStrRiyoDtFrom
        End Get
        Set(ByVal value As String)
            ppStrRiyoDtFrom = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用日（To）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyoDtTo</returns>
    ''' <remarks><para>作成情報：2015/09/09 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrRiyoDtTo() As String
        Get
            Return ppStrRiyoDtTo
        End Get
        Set(ByVal value As String)
            ppStrRiyoDtTo = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者名】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyoNm</returns>
    ''' <remarks><para>作成情報：2015/09/09 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrRiyoNm() As String
        Get
            Return ppStrRiyoNm
        End Get
        Set(ByVal value As String)
            ppStrRiyoNm = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者名カナ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyokana</returns>
    ''' <remarks><para>作成情報：2015/09/09 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrRiyokana() As String
        Get
            Return ppStrRiyokana
        End Get
        Set(ByVal value As String)
            ppStrRiyokana = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【催事名】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrSaijiNm</returns>
    ''' <remarks><para>作成情報：2015/09/09 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrSaijiNm() As String
        Get
            Return ppStrSaijiNm
        End Get
        Set(ByVal value As String)
            ppStrSaijiNm = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【予約NO】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrYoyakuNo</returns>
    ''' <remarks><para>作成情報：2015/09/09 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrYoyakuNo() As String
        Get
            Return ppStrYoyakuNo
        End Get
        Set(ByVal value As String)
            ppStrYoyakuNo = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【未完了フラグ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppBlnMikanryo</returns>
    ''' <remarks><para>作成情報：2015/09/09 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropBlnMikanryo() As Boolean
        Get
            Return ppBlnMikanryo
        End Get
        Set(ByVal value As Boolean)
            ppBlnMikanryo = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【予約一覧（シアター）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppvwYoyakuTheatre</returns>
    ''' <remarks><para>作成情報：2015/09/09 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropvwYoyakuTheatre() As FarPoint.Win.Spread.FpSpread
        Get
            Return ppvwYoyakuTheatre
        End Get
        Set(ByVal value As FarPoint.Win.Spread.FpSpread)
            ppvwYoyakuTheatre = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【予約一覧（スタジオ）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppvwYoyakuStudio</returns>
    ''' <remarks><para>作成情報：2015/09/09 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropvwYoyakuStudio() As FarPoint.Win.Spread.FpSpread
        Get
            Return ppvwYoyakuStudio
        End Get
        Set(ByVal value As FarPoint.Win.Spread.FpSpread)
            ppvwYoyakuStudio = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【予約データ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtYoyaku</returns>
    ''' <remarks><para>作成情報：2015/09/10 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtYoyaku() As DataTable
        Get
            Return ppDtYoyaku
        End Get
        Set(ByVal value As DataTable)
            ppDtYoyaku = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【選択確定ボタン】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppBtnDecision</returns>
    ''' <remarks><para>作成情報：2015/09/21 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropBtnDecision() As Button
        Get
            Return ppBtnDecision
        End Get
        Set(ByVal value As Button)
            ppBtnDecision = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【予約一覧フォーム】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppFrmYoyaku</returns>
    ''' <remarks><para>作成情報：2015/09/21 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropFrmYoyaku() As Form
        Get
            Return ppFrmYoyaku
        End Get
        Set(ByVal value As Form)
            ppFrmYoyaku = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【選択モード】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrSelectMode</returns>
    ''' <remarks><para>作成情報：2015/09/28 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrSelectMode() As String
        Get
            Return ppStrSelectMode
        End Get
        Set(ByVal value As String)
            ppStrSelectMode = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【遷移元戻し予約NO】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRtnYoyakuNo</returns>
    ''' <remarks><para>作成情報：2015/09/28 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrRtnYoyakuNo() As String
        Get
            Return ppStrRtnYoyakuNo
        End Get
        Set(ByVal value As String)
            ppStrRtnYoyakuNo = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【遷移元戻し催事名】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRtnSaijiNm</returns>
    ''' <remarks><para>作成情報：2015/09/28 h.endo
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrRtnSaijiNm() As String
        Get
            Return ppStrRtnSaijiNm
        End Get
        Set(ByVal value As String)
            ppStrRtnSaijiNm = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【チェックされた行番号】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppIntCheckRow</returns>
    ''' <remarks><para>作成情報：2015/12/01 y.ozawa
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

    ''' <summary>
    ''' プロパティセット【選択モード(施設区分) 】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrSelectShisetu</returns>
    ''' <remarks><para>作成情報：2015.12.15 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrSelectShisetu() As String
        Get
            Return ppStrSelectShisetu
        End Get
        Set(ByVal value As String)
            ppStrSelectShisetu = value
        End Set
    End Property

End Class
