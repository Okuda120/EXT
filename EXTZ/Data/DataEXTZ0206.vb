Public Class DataEXTZ0206

    Private ppBlnChangeFlg As Boolean
    Private ppStrFromName As String
    'Private ppIntFromKakutei As Integer              ' 2015.12.21 UPD h.hagiwara
    Private ppIntFromKakutei As Long                  ' 2015.12.21 UPD h.hagiwara
    Private ppStrFromNaiyo As String
    Private ppStrToName1 As String
    'Private ppIntToKakutei1 As Integer               ' 2015.12.21 UPD h.hagiwara
    Private ppIntToKakutei1 As Long                   ' 2015.12.21 UPD h.hagiwara
    Private ppStrToNaiyo1 As String
    Private ppStrToName2 As String
    'Private ppIntToKakutei2 As Integer               ' 2015.12.21 UPD h.hagiwara
    Private ppIntToKakutei2 As Long                   ' 2015.12.21 UPD h.hagiwara
    Private ppStrToNaiyo2 As String
    Private ppIntWariai1 As Integer
    Private ppIntWariai2 As Integer
    'Private ppIntSagaku As Integer                   ' 2015.12.21 UPD h.hagiwara
    Private ppIntSagaku As Long                       ' 2015.12.21 UPD h.hagiwara

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
    ''' プロパティセット【分割元:相手先名】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrFromName</returns>
    Public Property PropStrFromName()
        Get
            Return ppStrFromName
        End Get
        Set(ByVal value)
            ppStrFromName = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【分割元:確定額】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppIntFromKakutei</returns>
    Public Property PropIntFromKakutei()
        Get
            Return ppIntFromKakutei
        End Get
        Set(ByVal value)
            ppIntFromKakutei = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【分割元:請求内容】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrFromNaiyo</returns>
    Public Property PropStrFromNaiyo()
        Get
            Return ppStrFromNaiyo
        End Get
        Set(ByVal value)
            ppStrFromNaiyo = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【先1:相手先名】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrToName1</returns>
    Public Property PropStrToName1()
        Get
            Return ppStrToName1
        End Get
        Set(ByVal value)
            ppStrToName1 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【先1:確定額】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppIntToKakutei1</returns>
    Public Property PropIntToKakutei1()
        Get
            Return ppIntToKakutei1
        End Get
        Set(ByVal value)
            ppIntToKakutei1 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【先1:請求内容】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrToNaiyo1</returns>
    Public Property PropStrToNaiyo1()
        Get
            Return ppStrToNaiyo1
        End Get
        Set(ByVal value)
            ppStrToNaiyo1 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【先2:相手先名】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrToName2</returns>
    Public Property PropStrToName2()
        Get
            Return ppStrToName2
        End Get
        Set(ByVal value)
            ppStrToName2 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【先2:確定額】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppIntToKakutei2</returns>
    Public Property PropIntToKakutei2()
        Get
            Return ppIntToKakutei2
        End Get
        Set(ByVal value)
            ppIntToKakutei2 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【先2:請求内容】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrToNaiyo2</returns>
    Public Property PropStrToNaiyo2()
        Get
            Return ppStrToNaiyo2
        End Get
        Set(ByVal value)
            ppStrToNaiyo2 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【分配率１】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppIntWariai1</returns>
    Public Property PropIntWariai1()
        Get
            Return ppIntWariai1
        End Get
        Set(ByVal value)
            ppIntWariai1 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【分配率２】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppIntWariai2</returns>
    Public Property PropIntWariai2()
        Get
            Return ppIntWariai2
        End Get
        Set(ByVal value)
            ppIntWariai2 = value
        End Set
    End Property


    ''' <summary>
    ''' プロパティセット【差額】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppIntSagaku</returns>
    Public Property PropIntSagaku()
        Get
            Return ppIntSagaku
        End Get
        Set(ByVal value)
            ppIntSagaku = value
        End Set
    End Property
End Class
