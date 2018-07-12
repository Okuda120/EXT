Imports FarPoint.Win.Spread

Public Class DataEXTZ0212

    Private ppStrSelectRiyobi As String   '選択した利用日
    Private ppLstRiyobi As ArrayList      '利用日一覧
    Private ppDsTanakaBairitu As DataSet  '利用分類/料金/倍率
    Private ppStrShisetu As String        '施設区分
    Private ppBlnChangeFlg As Boolean     '変更フラグ
    Private ppIntCheckRow As Integer      'チェック行番号
    Private ppIntCheckCol As Integer      'チェック列番号
    Private fbRyokin As FpSpread          '料金一覧シート
    Private fbBairitu As FpSpread         '倍率一覧シート

    ''' <summary>
    ''' プロパティセット【参照ボタン押下行の利用日】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrSelectRiyobi</returns>
    Public Property PropStrSelectRiyobi()
        Get
            Return ppStrSelectRiyobi
        End Get
        Set(ByVal value)
            ppStrSelectRiyobi = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用日一覧】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppLstRiyobi</returns>
    Public Property PropLstRiyobi()
        Get
            Return ppLstRiyobi
        End Get
        Set(ByVal value)
            ppLstRiyobi = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用分類/料金/倍率】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDsTanakaBairitu</returns>
    Public Property PropDsTanakaBairitu()
        Get
            Return ppDsTanakaBairitu
        End Get
        Set(ByVal value)
            ppDsTanakaBairitu = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【施設区分】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDsBairitu</returns>
    Public Property PropStrShisetu()
        Get
            Return ppStrShisetu
        End Get
        Set(ByVal value)
            ppStrShisetu = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【変更フラグ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDsBairitu</returns>
    Public Property PropBlnChangeFlg()
        Get
            Return ppBlnChangeFlg
        End Get
        Set(ByVal value)
            ppBlnChangeFlg = value
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

    ''' <summary>
    ''' プロパティセット【チェック列番号】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppIntCheckIndex</returns>
    ''' <remarks><para>作成情報：2015/11/30 hayabuchi
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropIntCheckCol() As Integer
        Get
            Return ppIntCheckCol
        End Get
        Set(ByVal value As Integer)
            ppIntCheckCol = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【料金一覧シート】
    ''' </summary>
    ''' <value></value>
    ''' <returns>fbRyokin</returns>
    ''' <remarks><para>作成情報：2015/11/30 hayabuchi
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropRyokin() As FpSpread
        Get
            Return fbRyokin
        End Get
        Set(ByVal value As FpSpread)
            fbRyokin = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【倍率一覧シート】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppVwList</returns>
    ''' <remarks><para>作成情報：2015/11/30 hayabuchi
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropBairitu() As FpSpread
        Get
            Return fbBairitu
        End Get
        Set(ByVal value As FpSpread)
            fbBairitu = value
        End Set
    End Property

End Class
