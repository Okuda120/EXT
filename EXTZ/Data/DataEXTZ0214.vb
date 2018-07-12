
Public Class DataEXTZ0214

    Private ppStrSeikyusakiBusyoCD As String 'G請求先部署
    Private ppStrSeikyusyoTantoCD As String 'G請求書担当者コード
    Private ppStrSeikyuNaiyoCD As String 'G請求内容コード
    Private ppStrAitesakiCD As String   'EXAS相手先CD
    Private ppStrAitesakiNm As String      'EXAS相手先名
    Private ppTxtFilePathSeikyusakiBusyo As String      'path
    Private ppTxtFilePathSeikyuNaiyo As String      'path
    Private ppCmbSeikyusakiBusyo As ComboBox    'G請求先部署コンボボックス
    Private ppCmbSeikyuNaiyo As ComboBox    'G請求内容コンボボックス
    Private ppBlnGSeikyuFlg As Boolean      'グループ請求フラグ
    Private ppBlnGSeikyuCloseFlg As Boolean      'グループ請求クローズフラグ
    Private ppTblSeikyusakiBusyo As DataTable   'G請求先部署コンボボックス用データテーブル
    Private ppTblSeikyuNaiyo As DataTable     'G請求内容コンボボックス用データテーブル



    ''' <summary>
    ''' プロパティセット【G請求先部署コード】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrSeikyusakiBusyoCD</returns>
    ''' <remarks><para>作成情報：2016/01/18 y.morooka
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrSeikyusakiBusyoCD()
        Get
            Return ppStrSeikyusakiBusyoCD
        End Get
        Set(ByVal value)
            ppStrSeikyusakiBusyoCD = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【G請求書担当者コード】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrSeikyusyoTantoCD</returns>
    ''' <remarks><para>作成情報：2016/01/18 y.morooka
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrSeikyusyoTantoCD()
        Get
            Return ppStrSeikyusyoTantoCD
        End Get
        Set(ByVal value)
            ppStrSeikyusyoTantoCD = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【G請求内容コード】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrSeikyuNaiyoCD</returns>
    ''' <remarks><para>作成情報：2016/01/18 y.morooka
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrSeikyuNaiyoCD()
        Get
            Return ppStrSeikyuNaiyoCD
        End Get
        Set(ByVal value)
            ppStrSeikyuNaiyoCD = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【EXAS相手先CD】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppstrAitesakiCD</returns>
    Public Property PropStrAitesakiCD()
        Get
            Return ppStrAitesakiCD
        End Get
        Set(ByVal value)
            ppStrAitesakiCD = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【EXAS相手先名】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrAitesakiNm</returns>
    Public Property PropStrAitesakiNm()
        Get
            Return ppStrAitesakiNm
        End Get
        Set(ByVal value)
            ppStrAitesakiNm = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【G請求先部署CSVファイル：ファイルパス表示】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppTxtFilePathSeikyusakiBusyo</returns>
    ''' <remarks><para>作成情報：2016/01/20 y.morooka
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropTxtFilePathSeikyusakiBusyo() As String
        Get
            Return ppTxtFilePathSeikyusakiBusyo
        End Get
        Set(ByVal value As String)
            ppTxtFilePathSeikyusakiBusyo = value
        End Set
    End Property


    ''' <summary>
    ''' プロパティセット【G請求内容CSVファイル：ファイルパス表示】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppTxtFilePathSeikyuNaiyo</returns>
    ''' <remarks><para>作成情報：2016/01/20 y.morooka
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropTxtFilePathSeikyuNaiyo() As String
        Get
            Return ppTxtFilePathSeikyuNaiyo
        End Get
        Set(ByVal value As String)
            ppTxtFilePathSeikyuNaiyo = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【G請求先コンボボックス】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppcmbSeikyusakiBusyo</returns>
    ''' <remarks><para>作成情報：2016/01/19 y.morooka
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropCmbSeikyusakiBusyo() As ComboBox
        Get
            Return ppCmbSeikyusakiBusyo
        End Get
        Set(ByVal value As ComboBox)
            ppCmbSeikyusakiBusyo = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【G請求内容コンボボックス】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppCmbSeikyuNaiyo</returns>
    ''' <remarks><para>作成情報：2016/01/19 y.morooka
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropCmbSeikyuNaiyo() As ComboBox
        Get
            Return ppCmbSeikyuNaiyo
        End Get
        Set(ByVal value As ComboBox)
            ppCmbSeikyuNaiyo = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【グループ請求対応:出力判定フラグ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppsBlnGSeikyuFlg</returns>
    ''' <remarks><para>作成情報：2016/01/20 y.morooka
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropBlnGSeikyuFlg() As Boolean
        Get
            Return ppBlnGSeikyuFlg
        End Get
        Set(ByVal value As Boolean)
            ppBlnGSeikyuFlg = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【グループ請求対応:クローズ判定フラグ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppBlnGSeikyCloseFlg</returns>
    ''' <remarks><para>作成情報：2016/01/20 y.morooka
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropBlnGSeikyuCloseFlg() As Boolean
        Get
            Return ppBlnGSeikyuCloseFlg
        End Get
        Set(ByVal value As Boolean)
            ppBlnGSeikyuCloseFlg = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【G請求先コンボボックス設定用】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppTblSeikyusakiBusyo</returns>
    ''' <remarks><para>作成情報：2016/01/20 y.morooka
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropTblSeikyusakiBusyo As DataTable
        Get
            Return ppTblSeikyusakiBusyo
        End Get
        Set(ByVal value As DataTable)
            ppTblSeikyusakiBusyo = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【G請求内容コンボボックス設定用】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppTblSeikyuNaiyo</returns>
    ''' <remarks><para>作成情報：2016/01/20 y.morooka
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropTblSeikyuNaiyo As DataTable
        Get
            Return ppTblSeikyuNaiyo
        End Get
        Set(ByVal value As DataTable)
            ppTblSeikyuNaiyo = value
        End Set
    End Property

End Class
