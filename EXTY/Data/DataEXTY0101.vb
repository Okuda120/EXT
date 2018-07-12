Imports FarPoint.Win.Spread
Public Class DataEXTY0101

    '画面情報
    Private ppRdoTheater As RadioButton 'ラジオボタン（シアター）
    Private ppRdoStudio As RadioButton 'ラジオボタン（スタジオ）
    Private ppVwBillpay As FpSpread 'EXAS請求依頼スプレッドシート
    Private ppDtExasRequestData As DataTable 'EXAS請求依頼情報
    Private ppBlnCsvOutputFlg As Boolean 'CSV出力フラグ

    '検索条件
    Private ppStrBillNo_Search As String '検索条件：請求依頼番号
    Private ppstrReserveNo_Search As String '検索条件：予約NO
    Private ppStrSeq_Search As String '検索条件：予約NOシーケンス
    Private ppArySearchConditionList As ArrayList '検索条件格納用ArrayList

    'CSV出力用データテーブル
    Private ppDtOutputCsvData As DataTable 'EXAS請求依頼連携CSVデータ用データテーブル

    'データ
    Private ppIntCheckRow As Integer    'チェックされた行番号　2015.12.01 y.ozawa 単一チェック対応
    Private ppStrTantoNm As String      '担当者名              2016.08.12 m.hayabuchi 代行処理対応
    Private ppStrTantobusho As String   '所属部署名            2016.08.12 m.hayabuchi 代行処理対応

    ''' <summary>
    ''' プロパティセット【ラジオボタン（シアター）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppRdoTheater</returns>
    ''' <remarks><para>作成情報：2015/08/20 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropRdoTheater() As RadioButton
        Get
            Return ppRdoTheater
        End Get
        Set(ByVal value As RadioButton)
            ppRdoTheater = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【ラジオボタン（スタジオ）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppRdoStudio</returns>
    ''' <remarks><para>作成情報：2015/08/20 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropRdoStudio() As RadioButton
        Get
            Return ppRdoStudio
        End Get
        Set(ByVal value As RadioButton)
            ppRdoStudio = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【EXAS請求依頼スプレッドシート】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppVwBillpay</returns>
    ''' <remarks><para>作成情報：2015/08/21 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropVwBillpay() As FpSpread
        Get
            Return ppVwBillpay
        End Get
        Set(ByVal value As FpSpread)
            ppVwBillpay = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【EXAS請求依頼情報】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtExasRequestData</returns>
    ''' <remarks><para>作成情報：2015/08/21 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtExasRequestData() As DataTable
        Get
            Return ppDtExasRequestData
        End Get
        Set(ByVal value As DataTable)
            ppDtExasRequestData = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【検索条件：請求依頼番号】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrBillNo_Search</returns>
    ''' <remarks><para>作成情報：2015/08/24 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrBillNo_Search() As String
        Get
            Return ppStrBillNo_Search
        End Get
        Set(ByVal value As String)
            ppStrBillNo_Search = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【検索条件：予約NO】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppstrReserveNo_Search</returns>
    ''' <remarks><para>作成情報：2015/08/26 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropstrReserveNo_Search() As String
        Get
            Return ppstrReserveNo_Search
        End Get
        Set(ByVal value As String)
            ppstrReserveNo_Search = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【検索条件：予約NOシーケンス】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrSeq_Search</returns>
    ''' <remarks><para>作成情報：2015/08/26 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrSeq_Search() As String
        Get
            Return ppStrSeq_Search
        End Get
        Set(ByVal value As String)
            ppStrSeq_Search = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【検索条件格納用ArrayList】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppArySearchConditionList</returns>
    ''' <remarks><para>作成情報：2015/08/26 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropArySearchConditionList() As ArrayList
        Get
            Return ppArySearchConditionList
        End Get
        Set(ByVal value As ArrayList)
            ppArySearchConditionList = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【EXAS請求依頼連携CSVデータ用データテーブル】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtOutputCsvData</returns>
    ''' <remarks><para>作成情報：2015/08/25 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtOutputCsvData() As DataTable
        Get
            Return ppDtOutputCsvData
        End Get
        Set(ByVal value As DataTable)
            ppDtOutputCsvData = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【CSV出力フラグ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppBlnCsvOutputFlg</returns>
    ''' <remarks><para>作成情報：2015/08/25 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropBlnCsvOutputFlg() As Boolean
        Get
            Return ppBlnCsvOutputFlg
        End Get
        Set(ByVal value As Boolean)
            ppBlnCsvOutputFlg = value
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
    ''' プロパティセット【担当者名】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrTantoNm</returns>
    ''' <remarks><para>作成情報：2016/08/12 m.hayabuchi
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrTantoNm() As String
        Get
            Return ppStrTantoNm
        End Get
        Set(ByVal value As String)
            ppStrTantoNm = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【所属部署名】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrTantoBusho</returns>
    ''' <remarks><para>作成情報：2016/08/12 m.hayabuchi
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrTantoBusho() As String
        Get
            Return ppStrTantobusho
        End Get
        Set(ByVal value As String)
            ppStrTantobusho = value
        End Set
    End Property

End Class
