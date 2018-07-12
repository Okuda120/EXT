Public Class DataEXTZ0204

    Private ppBlnChangeFlg As Boolean
    Private ppDblTax As Double      'TAX
    Private ppDrBillReq As DataRow  '請求
    Private ppDtExasRiyoryo As DataTable  'EXASPRJ利用料
    Private ppDtExasFutai As DataTable  'EXASPRJ付帯
    Private ppStrTitle1 As String 'タイトル
    Private ppStrTitle2 As String 'タイトル
    Private ppStrSeikyuInputFlg As String

    ''' <summary>
    ''' プロパティセット【変更フラグ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppBlnChangeFlg</returns>
    ''' <remarks><para>作成情報：2015/09/01 k.machida
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
    ''' プロパティセット【消費税】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDblTax</returns>
    ''' <remarks><para>作成情報：2015/09/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDblTax()
        Get
            Return ppDblTax
        End Get
        Set(ByVal value)
            ppDblTax = value
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
    ''' プロパティセット【データテーブル:EXAS利用料】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtExasRiyoryo</returns>
    ''' <remarks><para>作成情報：2015/09/07 k.machida
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
    ''' <remarks><para>作成情報：2015/09/07 k.machida
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
    ''' プロパティセット【タイトル１】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrTitle1</returns>
    ''' <remarks><para>作成情報：2015/08/27 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property propStrTitle1()
        Get
            Return ppStrTitle1
        End Get
        Set(ByVal value)
            ppStrTitle1 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【タイトル２】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrTitle2</returns>
    ''' <remarks><para>作成情報：2015/08/27 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property propStrTitle2()
        Get
            Return ppStrTitle2
        End Get
        Set(ByVal value)
            ppStrTitle2 = value
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

End Class
