Imports FarPoint.Win.Spread

Public Class DataEXTM0202
    'ホームオブジェクト
    Private ppLblRiyo_cd As Label        '利用者番号
    Private ppTxtRiyo_kana As TextBox    '利用者カナ
    Private ppTxtRiyo_nm As TextBox      '利用者名
    Private ppTxtDaihyo_nm As TextBox    '代表者名
    Private ppTxtTel1 As TextBox         '電話番号1
    Private ppTxtTel2 As TextBox         '電話番号2
    Private ppTxtTel3 As TextBox         '電話番号3
    Private ppTxtNaisen As TextBox       '内線番号
    Private ppTxtFax1 As TextBox         'Fax番号1
    Private ppTxtFax2 As TextBox         'Fax番号2
    Private ppTxtFax3 As TextBox         'Fax番号3
    Private ppTxtYubin1 As TextBox       '郵便番号1
    Private ppTxtYubin2 As TextBox       '郵便番号2
    Private ppCmbTodo As ComboBox        '都道府県
    Private ppTxtShiku As TextBox        '市区町村
    Private ppTxtBanchi As TextBox       '番地
    Private ppTxtBuild As TextBox        'ビル名
    Private ppTxtCom As TextBox          'コメント
    Private ppPnlLevel As Panel          'レベル
    Private ppRdoTujyo As RadioButton    'レベル(通常)
    Private ppRdoChui As RadioButton     'レベル(要注意)
    Private ppRdoHuka As RadioButton     'レベル(利用不可)
    Private ppLblAite_cd As Label        '相手先コード
    Private ppLblAite_nm As Label        '相手先名

    Private ppVwList As FpSpread         '一覧シート

    '前画面からのパラメータ
    Private ppParamRiyoCd As String         '遷移元画面から引き渡される利用者番号
    Private ppParamRiyoNm As String         ' 遷移元画面から引き継がれる利用者名
    Private ppParamRiyoKana As String       ' 遷移元画面から引き継がれる利用者名カナ
    Private ppParamDaihyoNm As String       ' 遷移元画面から引き継がれる代表者名
    Private ppParamRiyoTel11 As String      ' 遷移元画面から引き継がれる利用者電話番号1
    Private ppParamRiyoTel12 As String      ' 遷移元画面から引き継がれる利用者電話番号2
    Private ppParamRiyoTel13 As String      ' 遷移元画面から引き継がれる利用者電話番号3
    Private ppParamRiyoNaisen As String     ' 遷移元画面から引き継がれる利用者内線番号
    Private ppParamRiyoFax11 As String      ' 遷移元画面から引き継がれる利用者FAX1
    Private ppParamRiyoFax12 As String      ' 遷移元画面から引き継がれる利用者FAX2
    Private ppParamRiyoFax13 As String      ' 遷移元画面から引き継がれる利用者FAX3
    Private ppParamRiyoYubin1 As String     ' 遷移元画面から引き継がれる利用者郵便番号1
    Private ppParamRiyoYubin2 As String     ' 遷移元画面から引き継がれる利用者郵便番号2
    Private ppParamRiyoTodo As String       ' 遷移元画面から引き継がれる利用者都道府県
    Private ppParamRiyoShiku As String      ' 遷移元画面から引き継がれる利用者市区町村
    Private ppParamRiyoBan As String        ' 遷移元画面から引き継がれる利用者番地
    Private ppParamRiyoBuild As String      ' 遷移元画面から引き継がれる利用者ビル名

    'データ
    Private ppDtRiyoshaMasta As DataTable   '利用者情報を格納するデータテーブル
    Private ppDtmSysDate As DateTime        'サーバー日付
    Private ppNewRiyo_cd As String         '新しく取得した利用者番号を格納

    ''' <summary>
    ''' プロパティセット【利用者番号】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppTxtRiyo_cd</returns>
    ''' <remarks><para>作成情報：2015/08/10 ozawa
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropLblRiyo_cd() As Label
        Get
            Return ppLblRiyo_cd
        End Get
        Set(ByVal value As Label)
            ppLblRiyo_cd = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者カナ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppTxtRiyo_kana</returns>
    ''' <remarks><para>作成情報：2015/08/17 ozawa
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropTxtRiyo_kana() As TextBox
        Get
            Return ppTxtRiyo_kana
        End Get
        Set(ByVal value As TextBox)
            ppTxtRiyo_kana = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者名】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppTxtRiyo_nm</returns>
    ''' <remarks><para>作成情報：2015/08/17 ozawa
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropTxtRiyo_nm() As TextBox
        Get
            Return ppTxtRiyo_nm
        End Get
        Set(ByVal value As TextBox)
            ppTxtRiyo_nm = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【代表者名】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppTxtDaihyo_nm</returns>
    ''' <remarks><para>作成情報：2015/08/26 ozawa
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropTxtDaihyo_nm() As TextBox
        Get
            Return ppTxtDaihyo_nm
        End Get
        Set(ByVal value As TextBox)
            ppTxtDaihyo_nm = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【電話番号①】
    ''' </summary>
    ''' <value></value>
    ''' <returns> ppTxtTel1</returns>
    ''' <remarks><para>作成情報：2015/08/17 ozawa
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropTxtTel1() As TextBox
        Get
            Return ppTxtTel1
        End Get
        Set(ByVal value As TextBox)
            ppTxtTel1 = value
        End Set
    End Property
    ''' <summary>
    ''' プロパティセット【電話番号②】
    ''' </summary>
    ''' <value></value>
    ''' <returns> ppTxtTel2</returns>
    ''' <remarks><para>作成情報：2015/08/17 ozawa
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropTxtTel2() As TextBox
        Get
            Return ppTxtTel2
        End Get
        Set(ByVal value As TextBox)
            ppTxtTel2 = value
        End Set
    End Property
    ''' <summary>
    ''' プロパティセット【電話番号③】
    ''' </summary>
    ''' <value></value>
    ''' <returns> ppTxtTel3</returns>
    ''' <remarks><para>作成情報：2015/08/17 ozawa
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropTxtTel3() As TextBox
        Get
            Return ppTxtTel3
        End Get
        Set(ByVal value As TextBox)
            ppTxtTel3 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【内線番号】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppTxtNaisen</returns>
    ''' <remarks><para>作成情報：2015/08/26 ozawa
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropTxtNaisen() As TextBox
        Get
            Return ppTxtNaisen
        End Get
        Set(ByVal value As TextBox)
            ppTxtNaisen = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【Fax番号1】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppTxtFax1</returns>
    ''' <remarks><para>作成情報：2015/08/26 ozawa
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropTxtFax1() As TextBox
        Get
            Return ppTxtFax1
        End Get
        Set(ByVal value As TextBox)
            ppTxtFax1 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【Fax番号2】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppTxtFax1</returns>
    ''' <remarks><para>作成情報：2015/08/26 ozawa
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropTxtFax2() As TextBox
        Get
            Return ppTxtFax2
        End Get
        Set(ByVal value As TextBox)
            ppTxtFax2 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【Fax番号3】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppTxtFax1</returns>
    ''' <remarks><para>作成情報：2015/08/26 ozawa
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropTxtFax3() As TextBox
        Get
            Return ppTxtFax3
        End Get
        Set(ByVal value As TextBox)
            ppTxtFax3 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【郵便番号1】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppTxtYubin1</returns>
    ''' <remarks><para>作成情報：2015/08/26 ozawa
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropTxtYubin1() As TextBox
        Get
            Return ppTxtYubin1
        End Get
        Set(ByVal value As TextBox)
            ppTxtYubin1 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【郵便番号2】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppTxtYubin1</returns>
    ''' <remarks><para>作成情報：2015/08/26 ozawa
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropTxtYubin2() As TextBox
        Get
            Return ppTxtYubin2
        End Get
        Set(ByVal value As TextBox)
            ppTxtYubin2 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【都道府県】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppCmbTodo</returns>
    ''' <remarks><para>作成情報：2015/08/26 ozawa
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropCmbTodo() As ComboBox
        Get
            Return ppCmbTodo
        End Get
        Set(ByVal value As ComboBox)
            ppCmbTodo = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【市区町村】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppTxtShiku</returns>
    ''' <remarks><para>作成情報：2015/08/26 ozawa
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropTxtShiku() As TextBox
        Get
            Return ppTxtShiku
        End Get
        Set(ByVal value As TextBox)
            ppTxtShiku = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【番地】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppTxtBanchi</returns>
    ''' <remarks><para>作成情報：2015/08/26 ozawa
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropTxtBanchi() As TextBox
        Get
            Return ppTxtBanchi
        End Get
        Set(ByVal value As TextBox)
            ppTxtBanchi = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【ビル名】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppTxtBuild</returns>
    ''' <remarks><para>作成情報：2015/08/26 ozawa
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropTxtBuild() As TextBox
        Get
            Return ppTxtBuild
        End Get
        Set(ByVal value As TextBox)
            ppTxtBuild = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【コメント】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppTxtCom</returns>
    ''' <remarks><para>作成情報：2015/08/26 ozawa
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropTxtCom() As TextBox
        Get
            Return ppTxtCom
        End Get
        Set(ByVal value As TextBox)
            ppTxtCom = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用レベル】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppRdoLevel</returns>
    ''' <remarks><para>作成情報：2015/08/10 ozawa
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropPnlLevel() As Panel
        Get
            Return ppPnlLevel
        End Get
        Set(ByVal value As Panel)
            ppPnlLevel = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用レベル（通常）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppRdoTujyo</returns>
    ''' <remarks><para>作成情報：2015/08/10 ozawa
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropRdoTujyo() As RadioButton
        Get
            Return ppRdoTujyo
        End Get
        Set(ByVal value As RadioButton)
            ppRdoTujyo = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用レベル（要注意）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppRdoChui</returns>
    ''' <remarks><para>作成情報：2015/08/10 ozawa
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropRdoChui() As RadioButton
        Get
            Return ppRdoChui
        End Get
        Set(ByVal value As RadioButton)
            ppRdoChui = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用レベル（利用不可）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppRdoHuka</returns>
    ''' <remarks><para>作成情報：2015/08/10 ozawa
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropRdoHuka() As RadioButton
        Get
            Return ppRdoHuka
        End Get
        Set(ByVal value As RadioButton)
            ppRdoHuka = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【相手先コード】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppTxtAite_cd</returns>
    ''' <remarks><para>作成情報：2015/08/17 ozawa
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropLblAite_cd() As Label
        Get
            Return ppLblAite_cd
        End Get
        Set(ByVal value As Label)
            ppLblAite_cd = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【相手先名】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppTxtAite_nm</returns>
    ''' <remarks><para>作成情報：2015/08/17 ozawa
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropLblAite_nm() As Label
        Get
            Return ppLblAite_nm
        End Get
        Set(ByVal value As Label)
            ppLblAite_nm = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【一覧シート】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppVwList</returns>
    ''' <remarks><para>作成情報：2012/06/15 matsuoka
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropVwList() As FpSpread
        Get
            Return ppVwList
        End Get
        Set(ByVal value As FpSpread)
            ppVwList = value
        End Set
    End Property


    ''' <summary>
    ''' プロパティセット【利用者番号（パラメータ）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppParamRiyoCd</returns>
    ''' <remarks><para>作成情報：2015/08/18 ozawa
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropParamRiyoCd() As String
        Get
            Return ppParamRiyoCd
        End Get
        Set(ByVal value As String)
            ppParamRiyoCd = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者情報を格納するデータテーブル】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtRiyoshaMasta</returns>
    ''' <remarks><para>作成情報：2015/08/26 ozawa
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtRiyoshaMasta() As DataTable
        Get
            Return ppDtRiyoshaMasta
        End Get
        Set(ByVal value As DataTable)
            ppDtRiyoshaMasta = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【サーバー日付】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtmSysDate</returns>
    ''' <remarks><para>作成情報：2015/08/25 ozawa
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtmSysDate() As DateTime
        Get
            Return ppDtmSysDate
        End Get
        Set(ByVal value As DateTime)
            ppDtmSysDate = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【新しく取得した利用者番号を格納する】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppNewRiyo_cd</returns>
    ''' <remarks><para>作成情報：2015/08/28 ozawa
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropNewRiyo_cd() As String
        Get
            Return ppNewRiyo_cd
        End Get
        Set(ByVal value As String)
            ppNewRiyo_cd = value
        End Set
    End Property

    ' 新規登録時の遷移元画面からの引き渡し項目
    ''' <summary>
    ''' プロパティセット【利用者名（パラメータ）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppParamRiyoNm</returns>
    ''' <remarks><para>作成情報：2015.10.14 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropParamRiyoNm() As String
        Get
            Return ppParamRiyoNm
        End Get
        Set(ByVal value As String)
            ppParamRiyoNm = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者名カナ（パラメータ）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppParamRiyoKana</returns>
    ''' <remarks><para>作成情報：2015.10.14 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropParamRiyoKana() As String
        Get
            Return ppParamRiyoKana
        End Get
        Set(ByVal value As String)
            ppParamRiyoKana = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【代表者名（パラメータ）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppParamDaihyoNm</returns>
    ''' <remarks><para>作成情報：2015.10.14 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropParamDaihyoNm() As String
        Get
            Return ppParamDaihyoNm
        End Get
        Set(ByVal value As String)
            ppParamDaihyoNm = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者電話番号1（パラメータ）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppParamRiyoTel11</returns>
    ''' <remarks><para>作成情報：2015.10.14 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropParamRiyoTel11() As String
        Get
            Return ppParamRiyoTel11
        End Get
        Set(ByVal value As String)
            ppParamRiyoTel11 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者電話番号2（パラメータ）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppParamRiyoTel12</returns>
    ''' <remarks><para>作成情報：2015.10.14 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropParamRiyoTel12() As String
        Get
            Return ppParamRiyoTel12
        End Get
        Set(ByVal value As String)
            ppParamRiyoTel12 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者電話番号3（パラメータ）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppParamRiyoTel13</returns>
    ''' <remarks><para>作成情報：2015.10.14 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropParamRiyoTel13() As String
        Get
            Return ppParamRiyoTel13
        End Get
        Set(ByVal value As String)
            ppParamRiyoTel13 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者内線番号（パラメータ）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppParamRiyoNaisen</returns>
    ''' <remarks><para>作成情報：2015.10.14 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropParamRiyoNaisen() As String
        Get
            Return ppParamRiyoNaisen
        End Get
        Set(ByVal value As String)
            ppParamRiyoNaisen = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者FAX1（パラメータ）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppParamRiyoFax11</returns>
    ''' <remarks><para>作成情報：2015.10.14 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropParamRiyoFax11() As String
        Get
            Return ppParamRiyoFax11
        End Get
        Set(ByVal value As String)
            ppParamRiyoFax11 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者FAX2（パラメータ）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppParamRiyoFax12</returns>
    ''' <remarks><para>作成情報：2015.10.14 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropParamRiyoFax12() As String
        Get
            Return ppParamRiyoFax12
        End Get
        Set(ByVal value As String)
            ppParamRiyoFax12 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者FAX3（パラメータ）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppParamRiyoFax13</returns>
    ''' <remarks><para>作成情報：2015.10.14 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropParamRiyoFax13() As String
        Get
            Return ppParamRiyoFax13
        End Get
        Set(ByVal value As String)
            ppParamRiyoFax13 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者郵便番号1（パラメータ）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppParamRiyoYubin1</returns>
    ''' <remarks><para>作成情報：2015.10.14 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropParamRiyoYubin1() As String
        Get
            Return ppParamRiyoYubin1
        End Get
        Set(ByVal value As String)
            ppParamRiyoYubin1 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者郵便番号2（パラメータ）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppParamRiyoYubin2</returns>
    ''' <remarks><para>作成情報：2015.10.14 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropParamRiyoYubin2() As String
        Get
            Return ppParamRiyoYubin2
        End Get
        Set(ByVal value As String)
            ppParamRiyoYubin2 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者都道府県（パラメータ）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppParamRiyoTodo</returns>
    ''' <remarks><para>作成情報：2015.10.14 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropParamRiyoTodo() As String
        Get
            Return ppParamRiyoTodo
        End Get
        Set(ByVal value As String)
            ppParamRiyoTodo = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者市区町村（パラメータ）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppParamRiyoShiku</returns>
    ''' <remarks><para>作成情報：2015.10.14 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropParamRiyoShiku() As String
        Get
            Return ppParamRiyoShiku
        End Get
        Set(ByVal value As String)
            ppParamRiyoShiku = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者番地（パラメータ）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppParamRiyoBan</returns>
    ''' <remarks><para>作成情報：2015.10.14 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropParamRiyoBan() As String
        Get
            Return ppParamRiyoBan
        End Get
        Set(ByVal value As String)
            ppParamRiyoBan = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者ビル名（パラメータ）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppParamRiyoBuild</returns>
    ''' <remarks><para>作成情報：2015.10.14 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropParamRiyoBuild() As String
        Get
            Return ppParamRiyoBuild
        End Get
        Set(ByVal value As String)
            ppParamRiyoBuild = value
        End Set
    End Property

End Class

