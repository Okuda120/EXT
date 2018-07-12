Imports FarPoint.Win.Spread

Public Class DataEXTM0104
    Private ppRdoshiki As RadioButton    '新規ラジオボタン
    Private ppRdosumi As RadioButton     '設定済みラジオボタン
    Private ppCmbKikan As ComboBox       '期間コンボボックス
    Private ppTxtKikanfromy As TextBox   '期間from年
    Private ppTxtKikanfromm As TextBox   '期間from月
    Private ppTxtKikantoy As TextBox     '期間to年
    Private ppTxtKikantom As TextBox     '期間to月
    Private ppVwBunrui As FpSpread      '分類マスタ
    Private ppVwRyokin As FpSpread      '料金マスタ
    Private ppVwBairitu As FpSpread     '倍率マスタ

    Private ppDtBunrui As DataTable      '分類テーブル
    Private ppDtRyokin As DataTable      '料金テーブル
    Private ppDtBairitu As DataTable     '倍率テーブル
    Private ppDtKikan As DataTable       '期間テーブル
    Private ppDtRyokinDsp As DataTable   '料金テーブル（表示設定用） 20151023 
    Private ppDtRyokinEscp As DataTable  '料金テーブル（退避用）     20151023

    Private ppStrBunruicd As String      'クリックされた分類コード
    Private ppHtKikan As Hashtable       '期間を格納するハッシュテーブル
    Private ppStrKikan As String         '検索用期間CD

    'insert用共通
    Private ppStrShisetu As String      '
    Private ppStrSort As String
    Private ppStrSts As String
    Private ppStrKikanfrom As String
    Private ppStrKikanto As String

    '分類マスタinsert用
    Private ppStrBunruicdinsert As String
    Private ppStrBunruinm As String

    '料金マスタinsert用
    Private ppStrRyokincd As String
    Private ppStrRyokinnm As String
    Private ppStrRyokinhour As String
    Private ppStrRyokin As String

    '倍率マスタinsert用
    Private ppStrBairitucd As String
    Private ppStrBairitunm As String
    Private ppStrBairitu As String

    '初期表示フラグ
    Private ppFlg As Boolean
    ' 共通
    Private ppDtSysDate As DateTime            ' サーバ日時


    ''' <summary>
    ''' プロパティセット【新規ラジオボタン】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppRdoshiki</returns>
    ''' <remarks><para>作成情報:2015/08/18 mori
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Property PropRdoshinki() As RadioButton
        Get
            Return ppRdoshiki
        End Get
        Set(ByVal value As RadioButton)
            ppRdoshiki = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【設定済みラジオボタン】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppRdosumi</returns>
    ''' <remarks><para>作成情報:2015/08/18 mori
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Property PropRdosumi() As RadioButton
        Get
            Return ppRdosumi
        End Get
        Set(ByVal value As RadioButton)
            ppRdosumi = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【期間データテーブル】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtKikan</returns>
    ''' <remarks><para>作成情報:2015/08/18 mori
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Property PropDtKikan() As DataTable
        Get
            Return ppDtKikan
        End Get
        Set(ByVal value As DataTable)
            ppDtKikan = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【期間コンボボックス】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppCmbKikan</returns>
    ''' <remarks><para>作成情報:2015/08/18 mori
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Property PropCmbKikan() As ComboBox
        Get
            Return ppCmbKikan
        End Get
        Set(ByVal value As ComboBox)
            ppCmbKikan = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【期間FROM年】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppTxtKikanfromy</returns>
    ''' <remarks><para>作成情報:2015/08/18 mori
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Property PropTxtKikanFromYear() As TextBox
        Get
            Return ppTxtKikanfromy
        End Get
        Set(ByVal value As TextBox)
            ppTxtKikanfromy = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【期間FROM月】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppTxtKikanfrommm</returns>
    ''' <remarks><para>作成情報:2015/08/18 mori
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Property PropTxtKikanFromMonth() As TextBox
        Get
            Return ppTxtKikanfromm
        End Get
        Set(ByVal value As TextBox)
            ppTxtKikanfromm = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【期間TO年】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppTxtKikantoy</returns>
    ''' <remarks><para>作成情報:2015/08/18 mori
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Property PropTxtKikanToYear() As TextBox
        Get
            Return ppTxtKikantoy
        End Get
        Set(ByVal value As TextBox)
            ppTxtKikantoy = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【期間TO月】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppTxtKikantom</returns>
    ''' <remarks><para>作成情報:2015/08/18 mori
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Property PropTxtkikanToMonth() As TextBox
        Get
            Return ppTxtKikantom
        End Get
        Set(ByVal value As TextBox)
            ppTxtKikantom = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【分類スブレット】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppVwBunrui</returns>
    ''' <remarks><para>作成情報:2015/08/18 mori
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Property PropVwBunrui() As FpSpread
        Get
            Return ppVwBunrui
        End Get
        Set(ByVal value As FpSpread)
            ppVwBunrui = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【料金スブレット】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppVwRyokin</returns>
    ''' <remarks><para>作成情報:2015/08/18 mori
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Property PropVwRyokin() As FpSpread
        Get
            Return ppVwRyokin
        End Get
        Set(ByVal value As FpSpread)
            ppVwRyokin = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【倍率スブレット】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppVwBairitu</returns>
    ''' <remarks><para>作成情報:2015/08/18 mori
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Property PropVwBairitu() As FpSpread
        Get
            Return ppVwBairitu
        End Get
        Set(ByVal value As FpSpread)
            ppVwBairitu = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【分類データテーブル】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtBunrui</returns>
    ''' <remarks><para>作成情報:2015/08/18 mori
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Property PropDtBunrui() As DataTable
        Get
            Return ppDtBunrui
        End Get
        Set(ByVal value As DataTable)
            ppDtBunrui = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【料金データテーブル】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtRyokin</returns>
    ''' <remarks><para>作成情報:2015/08/18 mori
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Property PropDtRyoukin() As DataTable
        Get
            Return ppDtRyokin
        End Get
        Set(ByVal value As DataTable)
            ppDtRyokin = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【倍率データテーブル】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtBairitu</returns>
    ''' <remarks><para>作成情報:2015/08/18 mori
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Property PropDtBairitu() As DataTable
        Get
            Return ppDtBairitu
        End Get
        Set(ByVal value As DataTable)
            ppDtBairitu = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【クリックされた分類CD】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrBunruicd</returns>
    ''' <remarks><para>作成情報:2015/08/18 mori
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Property PropStrBunruicd() As String
        Get
            Return ppStrBunruicd
        End Get
        Set(ByVal value As String)
            ppStrBunruicd = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【期間ハッシュテーブル】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppHtKikan</returns>
    ''' <remarks><para>作成情報:2015/08/18 mori
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Property PropHtKikan() As Hashtable
        Get
            Return ppHtKikan
        End Get
        Set(ByVal value As Hashtable)
            ppHtKikan = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【検索用期間】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrkikan</returns>
    ''' <remarks><para>作成情報:2015/08/18 mori
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Property PropStrKikan() As String
        Get
            Return ppStrKikan
        End Get
        Set(ByVal value As String)
            ppStrKikan = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【インサート用施設区分】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrShisetu</returns>
    ''' <remarks><para>作成情報:2015/08/18 mori
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Property PropStrShisetu() As String
        Get
            Return ppStrShisetu
        End Get
        Set(ByVal value As String)
            ppStrShisetu = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【インサート用ソート】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrSort</returns>
    ''' <remarks><para>作成情報:2015/08/18 mori
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Property PropStrSort() As String
        Get
            Return ppStrSort
        End Get
        Set(ByVal value As String)
            ppStrSort = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【インサート用期間FROM】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrKikanfrom</returns>
    ''' <remarks><para>作成情報:2015/08/18 mori
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Property PropStrKikanFrom() As String
        Get
            Return ppStrKikanfrom
        End Get
        Set(ByVal value As String)
            ppStrKikanfrom = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【インサート用期間TO】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrKikanto</returns>
    ''' <remarks><para>作成情報:2015/08/18 mori
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Property PropStrKikanTo() As String
        Get
            Return ppStrKikanto
        End Get
        Set(ByVal value As String)
            ppStrKikanto = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【インサート用ステータス】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrSts</returns>
    ''' <remarks><para>作成情報:2015/08/18 mori
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Property PropStrSts() As String
        Get
            Return ppStrSts
        End Get
        Set(ByVal value As String)
            ppStrSts = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【インサート用分類CD】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrBunruicdinsert</returns>
    ''' <remarks><para>作成情報:2015/08/18 mori
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Property PropStrBunruicdInsert() As String
        Get
            Return ppStrBunruicdinsert
        End Get
        Set(ByVal value As String)
            ppStrBunruicdinsert = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【インサート用分類名】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrBunruinm</returns>
    ''' <remarks><para>作成情報:2015/08/18 mori
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Property PropStrBunruinm() As String
        Get
            Return ppStrBunruinm
        End Get
        Set(ByVal value As String)
            ppStrBunruinm = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【インサート用料金CD】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRyokincd</returns>
    ''' <remarks><para>作成情報:2015/08/18 mori
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Property PropStrRyokincd() As String
        Get
            Return ppStrRyokincd
        End Get
        Set(ByVal value As String)
            ppStrRyokincd = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【インサート用料金名】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRyokinnm</returns>
    ''' <remarks><para>作成情報:2015/08/18 mori
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Property PropStrRyokinnm() As String
        Get
            Return ppStrRyokinnm
        End Get
        Set(ByVal value As String)
            ppStrRyokinnm = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【インサート用貸し時間】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRyokinhour</returns>
    ''' <remarks><para>作成情報:2015/08/18 mori
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Property PropStrRyokinhour() As String
        Get
            Return ppStrRyokinhour
        End Get
        Set(ByVal value As String)
            ppStrRyokinhour = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【インサート用料金】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRyokin</returns>
    ''' <remarks><para>作成情報:2015/08/18 mori
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Property PropStrRyokin() As String
        Get
            Return ppStrRyokin
        End Get
        Set(ByVal value As String)
            ppStrRyokin = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【インサート用倍率CD】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrBairitucd</returns>
    ''' <remarks><para>作成情報:2015/08/18 mori
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Property PropStrBairitucd() As String
        Get
            Return ppStrBairitucd
        End Get
        Set(ByVal value As String)
            ppStrBairitucd = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【期間データテーブル】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrBairitunm</returns>
    ''' <remarks><para>作成情報:2015/08/18 mori
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Property PropStrbairitunm() As String
        Get
            Return ppStrBairitunm
        End Get
        Set(ByVal value As String)
            ppStrBairitunm = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【インサート用倍率】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrBairitu</returns>
    ''' <remarks><para>作成情報:2015/08/18 mori
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Property propStrBairitu() As String
        Get
            Return ppStrBairitu
        End Get
        Set(ByVal value As String)
            ppStrBairitu = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【初期表示用フラグ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppFlg</returns>
    ''' <remarks><para>作成情報:2015/08/18 mori
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Property propFlg() As String
        Get
            Return ppFlg
        End Get
        Set(ByVal value As String)
            ppFlg = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【サーバ日時】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtAddDate</returns>
    ''' <remarks><para>作成情報：2015.10.16 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtSysDate() As DateTime
        Get
            Return ppDtSysDate
        End Get
        Set(ByVal value As DateTime)
            ppDtSysDate = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【料金データテーブル（表示用）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtRyokinDsp</returns>
    ''' <remarks><para>作成情報:2015/08/18 mori
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Property PropDtRyokinDsp() As DataTable
        Get
            Return ppDtRyokinDsp
        End Get
        Set(ByVal value As DataTable)
            ppDtRyokinDsp = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【料金データテーブル（退避用）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtRyokinEscp</returns>
    ''' <remarks><para>作成情報:2015/08/18 mori
    ''' <p>改訂情報：</p>
    ''' </para></remarks>
    Public Property PropDtRyokinEscp() As DataTable
        Get
            Return ppDtRyokinEscp
        End Get
        Set(ByVal value As DataTable)
            ppDtRyokinEscp = value
        End Set
    End Property

End Class

