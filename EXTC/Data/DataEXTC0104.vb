Public Class DataEXTC0104

    Private ppStrYoyakuNo As String                '予約No
    Private ppAryStrCancelDate As ArrayList        '利用日

    Private ppStrStatusNm As String
    Private ppStrCanUkeDt As String
    Private ppStrCanUkeUsercd As String
    Private ppStrYoyakuSts As String
    Private ppStrShisetuKbn As String
    Private ppStrStudioKbn As String
    Private ppStrShutsuenNm As String
    Private ppStrKashiKind As String
    Private ppStrOnkyoOpe As String
    Private ppStrRiyoshaCd As String
    Private ppStrRiyoNm As String
    Private ppStrRiyoKana As String
    Private ppStrSekininBusho_nm As String
    Private ppStrSekininNm As String
    Private ppStrSekininMail As String
    Private ppStrDaihyoNm As String
    Private ppStrRiyoTel11 As String
    Private ppStrRiyoTel12 As String
    Private ppStrRiyoTel13 As String
    Private ppStrRiyoTel21 As String
    Private ppStrRiyoTel22 As String
    Private ppStrRiyoTel23 As String
    Private ppStrRiyoNaisen As String
    Private ppStrRiyoFax11 As String
    Private ppStrRiyoFax12 As String
    Private ppStrRiyoFax13 As String
    Private ppStrRiyoYubin1 As String
    Private ppStrRiyoYubin2 As String
    Private ppStrRiyoTodo As String
    Private ppStrRiyoShiku As String
    Private ppStrRiyoBan As String
    Private ppStrRiyoBuild As String
    Private ppStrRiyoLvl As String
    Private ppStrAiteCd As String
    Private ppStrAiteNm As String
    Private ppStrBiko As String
    Private ppStrOnkyoNm As String
    Private ppStrOnkyoTantoNm As String
    Private ppStrOnkyoTel11 As String
    Private ppStrOnkyoTel12 As String
    Private ppStrOnkyoTel13 As String
    Private ppStrOnkyoNaisen As String
    Private ppStrOnkyoFax11 As String
    Private ppStrOnkyoFax12 As String
    Private ppStrOnkyoFax13 As String
    Private ppStrOnkyoMail As String

    Private ppStrAddDt As String
    Private ppStrAddUserCd As String
    Private ppStrAddUserNm As String
    Private ppStrUpDt As String
    Private ppStrUpUserCd As String
    Private ppStrUpUserNm As String

    Private ppListRiyobi As ArrayList


    ''' <summary>
    ''' プロパティセット【予約No】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrCancelNo</returns>
    ''' <remarks><para>作成情報：2015/08/20 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrYoyakuNo()
        Get
            Return ppStrYoyakuNo
        End Get
        Set(ByVal value)
            ppStrYoyakuNo = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用日リスト(ArrayList Array{String,String})】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppAryStrCancelDate</returns>
    ''' <remarks><para>作成情報：2015/08/20 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropAryStrCancelDate()
        Get
            Return ppAryStrCancelDate
        End Get
        Set(ByVal value)
            ppAryStrCancelDate = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【ステータス名】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrStatusNm</returns>
    ''' <remarks><para>作成情報：2015/09/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrStatusNm()
        Get
            Return ppStrStatusNm
        End Get
        Set(ByVal value)
            ppStrStatusNm = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【受付日】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrCanUkeDt</returns>
    ''' <remarks><para>作成情報：2015/09/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrCanUkeDt()
        Get
            Return ppStrCanUkeDt
        End Get
        Set(ByVal value)
            ppStrCanUkeDt = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【受付ユーザCD】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrCanUkeUsercd</returns>
    ''' <remarks><para>作成情報：2015/09/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrCanUkeUsercd()
        Get
            Return ppStrCanUkeUsercd
        End Get
        Set(ByVal value)
            ppStrCanUkeUsercd = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【ステータスCD】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrYoyakuSts</returns>
    ''' <remarks><para>作成情報：2015/09/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrYoyakuSts()
        Get
            Return ppStrYoyakuSts
        End Get
        Set(ByVal value)
            ppStrYoyakuSts = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【施設区分】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrShisetuKbn</returns>
    ''' <remarks><para>作成情報：2015/09/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrShisetuKbn()
        Get
            Return ppStrShisetuKbn
        End Get
        Set(ByVal value)
            ppStrShisetuKbn = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【スタジオ区分】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrStudioKbn</returns>
    ''' <remarks><para>作成情報：2015/09/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrStudioKbn()
        Get
            Return ppStrStudioKbn
        End Get
        Set(ByVal value)
            ppStrStudioKbn = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【出演名】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrShutsuenNm</returns>
    ''' <remarks><para>作成情報：2015/09/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrShutsuenNm()
        Get
            Return ppStrShutsuenNm
        End Get
        Set(ByVal value)
            ppStrShutsuenNm = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【貸出種別】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrKashiKind</returns>
    ''' <remarks><para>作成情報：2015/09/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrKashiKind()
        Get
            Return ppStrKashiKind
        End Get
        Set(ByVal value)
            ppStrKashiKind = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【音響オペレーター】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrOnkyoOpe</returns>ppStrKashiKind
    Public Property PropStrOnkyoOpe()
        Get
            Return ppStrOnkyoOpe
        End Get
        Set(ByVal value)
            ppStrOnkyoOpe = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者CD】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyoshaCd</returns>
    ''' <remarks><para>作成情報：2015/09/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrRiyoshaCd()
        Get
            Return ppStrRiyoshaCd
        End Get
        Set(ByVal value)
            ppStrRiyoshaCd = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者名】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyoNm</returns>
    ''' <remarks><para>作成情報：2015/09/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrRiyoNm()
        Get
            Return ppStrRiyoNm
        End Get
        Set(ByVal value)
            ppStrRiyoNm = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者名カナ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyoKana</returns>
    ''' <remarks><para>作成情報：2015/09/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrRiyoKana()
        Get
            Return ppStrRiyoKana
        End Get
        Set(ByVal value)
            ppStrRiyoKana = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【責任者部署名】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrSekininBusho_nm</returns>
    ''' <remarks><para>作成情報：2015/09/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrSekininBushoNm()
        Get
            Return ppStrSekininBusho_nm
        End Get
        Set(ByVal value)
            ppStrSekininBusho_nm = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【責任者名】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrSekininNm</returns>
    ''' <remarks><para>作成情報：2015/09/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrSekininNm()
        Get
            Return ppStrSekininNm
        End Get
        Set(ByVal value)
            ppStrSekininNm = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【責任者メールアドレス】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrSekininMail</returns>
    ''' <remarks><para>作成情報：2015/09/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrSekininMail()
        Get
            Return ppStrSekininMail
        End Get
        Set(ByVal value)
            ppStrSekininMail = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【代表者名】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrDaihyoNm</returns>
    ''' <remarks><para>作成情報：2015/09/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrDaihyoNm()
        Get
            Return ppStrDaihyoNm
        End Get
        Set(ByVal value)
            ppStrDaihyoNm = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者TEL】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyoTel11</returns>
    ''' <remarks><para>作成情報：2015/09/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrRiyoTel11()
        Get
            Return ppStrRiyoTel11
        End Get
        Set(ByVal value)
            ppStrRiyoTel11 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者TEL】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyoTel12</returns>
    ''' <remarks><para>作成情報：2015/09/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrRiyoTel12()
        Get
            Return ppStrRiyoTel12
        End Get
        Set(ByVal value)
            ppStrRiyoTel12 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者TEL】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyoTel13</returns>
    ''' <remarks><para>作成情報：2015/09/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrRiyoTel13()
        Get
            Return ppStrRiyoTel13
        End Get
        Set(ByVal value)
            ppStrRiyoTel13 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者携帯】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyoTel21</returns>
    ''' <remarks><para>作成情報：2015/09/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrRiyoTel21()
        Get
            Return ppStrRiyoTel21
        End Get
        Set(ByVal value)
            ppStrRiyoTel21 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者携帯】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyoTel22</returns>
    ''' <remarks><para>作成情報：2015/09/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrRiyoTel22()
        Get
            Return ppStrRiyoTel22
        End Get
        Set(ByVal value)
            ppStrRiyoTel22 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者携帯】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyoTel23</returns>
    ''' <remarks><para>作成情報：2015/09/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrRiyoTel23()
        Get
            Return ppStrRiyoTel23
        End Get
        Set(ByVal value)
            ppStrRiyoTel23 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者内戦】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyoNaisen</returns>
    ''' <remarks><para>作成情報：2015/09/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrRiyoNaisen()
        Get
            Return ppStrRiyoNaisen
        End Get
        Set(ByVal value)
            ppStrRiyoNaisen = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者FAX】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyoFax11</returns>
    ''' <remarks><para>作成情報：2015/09/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrRiyoFax11()
        Get
            Return ppStrRiyoFax11
        End Get
        Set(ByVal value)
            ppStrRiyoFax11 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者FAX】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyoFax12</returns>
    ''' <remarks><para>作成情報：2015/09/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrRiyoFax12()
        Get
            Return ppStrRiyoFax12
        End Get
        Set(ByVal value)
            ppStrRiyoFax12 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者FAX】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyoFax13</returns>
    ''' <remarks><para>作成情報：2015/09/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrRiyoFax13()
        Get
            Return ppStrRiyoFax13
        End Get
        Set(ByVal value)
            ppStrRiyoFax13 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【郵便番号】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyoYubin1</returns>
    ''' <remarks><para>作成情報：2015/09/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrRiyoYubin1()
        Get
            Return ppStrRiyoYubin1
        End Get
        Set(ByVal value)
            ppStrRiyoYubin1 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【郵便番号】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyoYubin2</returns>
    ''' <remarks><para>作成情報：2015/09/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrRiyoYubin2()
        Get
            Return ppStrRiyoYubin2
        End Get
        Set(ByVal value)
            ppStrRiyoYubin2 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【都道府県】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyoTodo</returns>
    ''' <remarks><para>作成情報：2015/09/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrRiyoTodo()
        Get
            Return ppStrRiyoTodo
        End Get
        Set(ByVal value)
            ppStrRiyoTodo = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【市区町村名】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyoShiku</returns>
    ''' <remarks><para>作成情報：2015/09/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrRiyoShiku()
        Get
            Return ppStrRiyoShiku
        End Get
        Set(ByVal value)
            ppStrRiyoShiku = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【番地】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyoBan</returns>
    ''' <remarks><para>作成情報：2015/09/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrRiyoBan()
        Get
            Return ppStrRiyoBan
        End Get
        Set(ByVal value)
            ppStrRiyoBan = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【ビル名】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyoBuild</returns>
    ''' <remarks><para>作成情報：2015/09/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrRiyoBuild()
        Get
            Return ppStrRiyoBuild
        End Get
        Set(ByVal value)
            ppStrRiyoBuild = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者LV】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyoLvl</returns>
    ''' <remarks><para>作成情報：2015/09/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrRiyoLvl()
        Get
            Return ppStrRiyoLvl
        End Get
        Set(ByVal value)
            ppStrRiyoLvl = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【相手先CD】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrAiteCd</returns>
    ''' <remarks><para>作成情報：2015/09/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrAiteCd()
        Get
            Return ppStrAiteCd
        End Get
        Set(ByVal value)
            ppStrAiteCd = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【相手先名】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrAiteNm</returns>
    ''' <remarks><para>作成情報：2015/09/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrAiteNm()
        Get
            Return ppStrAiteNm
        End Get
        Set(ByVal value)
            ppStrAiteNm = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【特記事項】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrBiko</returns>
    ''' <remarks><para>作成情報：2015/09/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrBiko()
        Get
            Return ppStrBiko
        End Get
        Set(ByVal value)
            ppStrBiko = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【音響会社名】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrOnkyoNm</returns>
    ''' <remarks><para>作成情報：2015/08/26 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrOnkyoNm()
        Get
            Return ppStrOnkyoNm
        End Get
        Set(ByVal value)
            ppStrOnkyoNm = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【音響会社担当者名】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrOnkyoTantoNm</returns>
    ''' <remarks><para>作成情報：2015/08/26 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrOnkyoTantoNm()
        Get
            Return ppStrOnkyoTantoNm
        End Get
        Set(ByVal value)
            ppStrOnkyoTantoNm = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【音響会社電話番号1-1】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrOnkyoTel11</returns>
    ''' <remarks><para>作成情報：2015/08/26 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrOnkyoTel11()
        Get
            Return ppStrOnkyoTel11
        End Get
        Set(ByVal value)
            ppStrOnkyoTel11 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【音響会社電話番号1-2】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrOnkyoTel12</returns>
    ''' <remarks><para>作成情報：2015/08/26 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrOnkyoTel12()
        Get
            Return ppStrOnkyoTel12
        End Get
        Set(ByVal value)
            ppStrOnkyoTel12 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【音響会社電話番号1-3】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrOnkyoTel13</returns>
    ''' <remarks><para>作成情報：2015/08/26 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrOnkyoTel13()
        Get
            Return ppStrOnkyoTel13
        End Get
        Set(ByVal value)
            ppStrOnkyoTel13 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【音響会社内線番号】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrOnkyoNaisen</returns>
    ''' <remarks><para>作成情報：2015/08/26 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrOnkyoNaisen()
        Get
            Return ppStrOnkyoNaisen
        End Get
        Set(ByVal value)
            ppStrOnkyoNaisen = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【音響会社FAX1-1】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrOnkyoFax11</returns>
    ''' <remarks><para>作成情報：2015/08/26 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrOnkyoFax11()
        Get
            Return ppStrOnkyoFax11
        End Get
        Set(ByVal value)
            ppStrOnkyoFax11 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【音響会社FAX1-2】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrOnkyoFax12</returns>
    ''' <remarks><para>作成情報：2015/08/26 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrOnkyoFax12()
        Get
            Return ppStrOnkyoFax12
        End Get
        Set(ByVal value)
            ppStrOnkyoFax12 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【音響会社FAX1-3】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrOnkyoFax13</returns>
    ''' <remarks><para>作成情報：2015/08/26 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrOnkyoFax13()
        Get
            Return ppStrOnkyoFax13
        End Get
        Set(ByVal value)
            ppStrOnkyoFax13 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【音響会社メールアドレス】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrOnkyoMail</returns>
    ''' <remarks><para>作成情報：2015/08/26 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrOnkyoMail()
        Get
            Return ppStrOnkyoMail
        End Get
        Set(ByVal value)
            ppStrOnkyoMail = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【更新日時】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrAddDt</returns>
    ''' <remarks><para>作成情報：2015/09/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrAddDt()
        Get
            Return ppStrAddDt
        End Get
        Set(ByVal value)
            ppStrAddDt = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【更新者ID】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrAddUserCd</returns>
    ''' <remarks><para>作成情報：2015/09/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrAddUserCd()
        Get
            Return ppStrAddUserCd
        End Get
        Set(ByVal value)
            ppStrAddUserCd = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【更新者名】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrAddUserNm</returns>
    ''' <remarks><para>作成情報：2015/09/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrAddUserNm()
        Get
            Return ppStrAddUserNm
        End Get
        Set(ByVal value)
            ppStrAddUserNm = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【作成日時】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrUpDt</returns>
    ''' <remarks><para>作成情報：2015/09/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrUpDt()
        Get
            Return ppStrUpDt
        End Get
        Set(ByVal value)
            ppStrUpDt = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【作成者ID】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrUpUserCd</returns>
    ''' <remarks><para>作成情報：2015/09/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrUpUserCd()
        Get
            Return ppStrUpUserCd
        End Get
        Set(ByVal value)
            ppStrUpUserCd = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【作成者名】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrUpUserNm</returns>
    ''' <remarks><para>作成情報：2015/09/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrUpUserNm()
        Get
            Return ppStrUpUserNm
        End Get
        Set(ByVal value)
            ppStrUpUserNm = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用日一覧】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppListRiyobi</returns>
    ''' <remarks><para>作成情報：2015/09/07 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropListRiyobi()
        Get
            Return ppListRiyobi
        End Get
        Set(ByVal value)
            ppListRiyobi = value
        End Set
    End Property

End Class
