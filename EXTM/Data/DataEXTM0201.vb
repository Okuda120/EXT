Imports FarPoint.Win.Spread

Public Class DataEXTM0201
    'フォームオブジェクト（検索条件）
    Private ppTxtRiyo_cd As TextBox    '利用者番号
    Private ppTxtRiyo_kana As TextBox  '利用者カナ
    Private ppTxtRiyo_nm As TextBox    '利用者名
    Private ppTxtTel1 As TextBox       '電話番号1
    Private ppTxtTel2 As TextBox       '電話番号2
    Private ppTxtTel3 As TextBox       '電話番号3
    Private ppTxtAite_cd As TextBox    '相手先コード
    Private ppTxtAite_nm As TextBox    '相手先名
    Private ppPnlLevel As Panel        'レベル
    Private ppRdoTujyo As RadioButton  'レベル(通常)
    Private ppRdoChui As RadioButton   'レベル(要注意)
    Private ppRdoHuka As RadioButton   'レベル(利用不可)
    Private ppVwList As FpSpread       '一覧シート
    Private ppRiyo_kana As String      '利用者カナ  2015.10.05 マージによる追加 systemthink
    Private ppInitDspFlg As String     ' 検索表示フラグ 0:表示なし 1:表示あり      2015.12.03 ADD h.hagiwara

    'データ
    Private ppIndex As Integer                      '列インデックス
    Private ppCheckIndex As Integer                 'チェックされた行番号　2015.11.30 単一チェック対応　y.ozawa

    Private ppParamValue As String                  '呼び出し元画面から渡ってくる値

    ' 引き渡しパラメータ
    Private ppParamRiyoCd As String         ' 遷移元画面に引き渡す利用者番号
    Private ppParamRiyoNm As String         ' 遷移元画面に引き渡す利用者名
    Private ppParamRiyoKana As String       ' 遷移元画面に引き渡す利用者名カナ
    Private ppParamDaihyoNm As String       ' 遷移元画面に引き渡す代表者名
    Private ppParamRiyoTel11 As String      ' 遷移元画面に引き渡す利用者電話番号1
    Private ppParamRiyoTel12 As String      ' 遷移元画面に引き渡す利用者電話番号2
    Private ppParamRiyoTel13 As String      ' 遷移元画面に引き渡す利用者電話番号3
    Private ppParamRiyoNaisen As String     ' 遷移元画面に引き渡す利用者内線番号
    Private ppParamRiyoFax11 As String      ' 遷移元画面に引き渡す利用者FAX1
    Private ppParamRiyoFax12 As String      ' 遷移元画面に引き渡す利用者FAX2
    Private ppParamRiyoFax13 As String      ' 遷移元画面に引き渡す利用者FAX3
    Private ppParamRiyoYubin1 As String     ' 遷移元画面に引き渡す利用者郵便番号1
    Private ppParamRiyoYubin2 As String     ' 遷移元画面に引き渡す利用者郵便番号2
    Private ppParamRiyoTodo As String       ' 遷移元画面に引き渡す利用者都道府県
    Private ppParamRiyoShiku As String      ' 遷移元画面に引き渡す利用者市区町村
    Private ppParamRiyoBan As String        ' 遷移元画面に引き渡す利用者番地
    Private ppParamRiyoBuild As String      ' 遷移元画面に引き渡す利用者ビル名
    Private ppParamRiyoLvl As String        ' 遷移元画面に引き渡す利用者レベル
    Private ppParamAiteCd As String         ' 遷移元画面に引き渡す相手先コード
    Private ppParamAiteNm As String         ' 遷移元画面に引き渡す相手先名

    ''' <summary>
    ''' プロパティセット【利用者番号】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppTxtRiyo_cd</returns>
    ''' <remarks><para>作成情報：2015/08/10 ozawa
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropTxtRiyo_cd() As TextBox
        Get
            Return ppTxtRiyo_cd
        End Get
        Set(ByVal value As TextBox)
            ppTxtRiyo_cd = value
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
    ''' プロパティセット【相手先コード】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppTxtAite_cd</returns>
    ''' <remarks><para>作成情報：2015/08/17 ozawa
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropTxtAite_cd() As TextBox
        Get
            Return ppTxtAite_cd
        End Get
        Set(ByVal value As TextBox)
            ppTxtAite_cd = value
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
    Public Property PropTxtAite_nm() As TextBox
        Get
            Return ppTxtAite_nm
        End Get
        Set(ByVal value As TextBox)
            ppTxtAite_nm = value
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
    ''' プロパティセット【列インデックス】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppIndex</returns>
    ''' <remarks><para>作成情報：2015/08/17 ozawa
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropIndex() As Integer
        Get
            Return ppIndex
        End Get
        Set(ByVal value As Integer)
            ppIndex = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【チェックされた行番号】
    ''' </summary>
    ''' <value></value>
    ''' <returns>CheckIndex</returns>
    ''' <remarks><para>作成情報：2015/11/30 y.ozawa
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropCheckIndex() As Integer
        Get
            Return ppCheckIndex
        End Get
        Set(ByVal value As Integer)
            ppCheckIndex = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【呼び出し元画面から渡ってくるパラメータ値】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppParamValue</returns>
    ''' <remarks><para>作成情報：2015/08/19 ozawa
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropParamValue() As String
        Get
            Return ppParamValue
        End Get
        Set(ByVal value As String)
            ppParamValue = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【呼び出し元画面から渡ってくるパラメータ値】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppParamValue</returns>
    ''' <remarks><para>作成情報：2015/10/05 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrRiyoKana() As String
        Get
            Return ppRiyo_kana
        End Get
        Set(ByVal value As String)
            ppRiyo_kana = value
        End Set
    End Property


    ' 遷移元引き渡し項目
    ''' <summary>
    ''' プロパティセット【利用者番号（パラメータ）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppParamRiyoCd</returns>
    ''' <remarks><para>作成情報：2015.10.14 h.hagiwara
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

    ''' <summary>
    ''' プロパティセット【利用者レベル（パラメータ）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppParamRiyoLvl</returns>
    ''' <remarks><para>作成情報：2015.10.14 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropParamRiyoLvl() As String
        Get
            Return ppParamRiyoLvl
        End Get
        Set(ByVal value As String)
            ppParamRiyoLvl = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【相手先コード（パラメータ）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppParamAiteCd</returns>
    ''' <remarks><para>作成情報：2015.10.14 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropParamAiteCd() As String
        Get
            Return ppParamAiteCd
        End Get
        Set(ByVal value As String)
            ppParamAiteCd = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【相手先名（パラメータ）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppParamAiteNm</returns>
    ''' <remarks><para>作成情報：2015.10.14 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropParamAiteNm() As String
        Get
            Return ppParamAiteNm
        End Get
        Set(ByVal value As String)
            ppParamAiteNm = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【検索表示フラグ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppInitDspFlg</returns>
    ''' <remarks><para>作成情報：2015.12.03 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropInitDspFlg() As String
        Get
            Return ppInitDspFlg
        End Get
        Set(ByVal value As String)
            ppInitDspFlg = value
        End Set
    End Property

End Class
