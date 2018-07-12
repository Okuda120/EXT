''' <summary>
''' DataEXTB0102
''' </summary>
''' <remarks>予約、正式予約（シアター）画面データクラス
''' <para>作成情報：2015/08/04 k.machida
''' <p>改訂情報:</p>
''' </para></remarks>
Public Class DataEXTB0102

    'パラメータ変数
    Private ppStrYoyakuNo As String                '予約No
    Private ppStrStatusNm As String                'ステータスNm
    Private ppStrKariukeDt As String
    Private ppStrUkeDt As String
    Private ppStrKariUsercd As String
    Private ppStrKakuteiDt As String
    Private ppStrKakuUsercd As String
    Private ppStrYoyakuSts As String
    Private ppStrShisetuKbn As String
    Private ppStrStudioKbn As String
    Private ppStrSaijiNm As String
    Private ppStrShutsuenNm As String
    Private ppStrKashiKind As String
    Private ppStrRiyoType As String
    Private ppStrDrinkFlg As String
    Private ppStrSaijiBunrui As String
    Private ppStrTeiin As String
    Private ppStrRiyoshaCd As String
    Private ppStrRiyoNm As String
    Private ppStrRiyoKana As String
    Private ppStrSekininBushoNm As String
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
    Private ppStrSendKbn As String
    Private ppStrSendSts As String
    Private ppStrSendDt As String
    Private ppStrHensoDt As String
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
    Private ppStrTotalRiyoKin As String
    Private ppStrRiyoCom As String
    Private ppStrTicketEnterKbn As String
    Private ppStrTicketDrinkKbn As String
    Private ppStrHpKeisai As String
    Private ppStrJohoKokaiDt As String
    Private ppStrJohoKokaiTime As String
    Private ppStrKokaiDt As String
    Private ppStrKokaiTime As String
    Private ppStrFinputSts As String

    Private ppStrAddDt As String
    Private ppStrAddUserCd As String
    Private ppStrAddUserNm As String
    Private ppStrUpDt As String
    Private ppStrUpUserCd As String
    Private ppStrUpUserNm As String

    Private ppListRiyobi As ArrayList
    Private ppAryStrRiyoDate As ArrayList
    Private ppStrCancelNo As String 'キャンセル画面より
    'Private ppIntTtl As Integer    '合計                    ' 2015.12.21 UPD h.hagiwara
    Private ppIntTtl As Long        '合計                    ' 2015.12.21 UPD h.hagiwara
    Private ppDblTax As Double      'TAX
    '各種データセット
    Private ppDsBillReq As DataSet '請求/入金/EXAS入金
    Private ppHtExasPro As Hashtable 'EXASプロジェクト設定
    Private ppHtExasProFutai As Hashtable 'EXASプロジェクト設定付帯設備
    Private ppDtExasNyukin As DataTable 'EXAS入金情報
    Private ppDsApproval As DataSet '承認依頼／記録
    Private ppDsTimeSch As DataSet 'タイムスケジュール
    Private ppDsTicket As DataSet  'Ticket
    Private ppHtFutai As Hashtable '付帯明細
    Private ppHtFutaiDetail As Hashtable '付帯明細
    Private ppDsCashElectro As DataSet '現金／電子

    ''' <summary>
    ''' プロパティセット【予約No】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrYoyakuNo</returns>
    ''' <remarks><para>作成情報：2015/08/10 k.machida
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
    ''' プロパティセット【STATUS_NAME】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrStatusNm</returns>
    ''' <remarks><para>作成情報：2015/08/12 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropptrStatusNm()
        Get
            Return ppStrStatusNm
        End Get
        Set(ByVal value)
            ppStrStatusNm = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【仮予約受付日】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrKariukeDt</returns>
    ''' <remarks><para>作成情報：2015/08/10 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrKariukeDt()
        Get
            Return ppStrKariukeDt
        End Get
        Set(ByVal value)
            ppStrKariukeDt = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【受付日】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrUkeDt</returns>
    ''' <remarks><para>作成情報：2015/08/20 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrUkeDt()
        Get
            Return ppStrUkeDt
        End Get
        Set(ByVal value)
            ppStrUkeDt = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【仮予約受付ﾕｰｻﾞCD】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrKariUsercd</returns>ppStrKariukeDt
    Public Property PropStrKariUsercd()
        Get
            Return ppStrKariUsercd
        End Get
        Set(ByVal value)
            ppStrKariUsercd = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【予約確定日】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrKakuteiDt</returns>ppStrKariUsercd
    Public Property PropStrKakuteiDt()
        Get
            Return ppStrKakuteiDt
        End Get
        Set(ByVal value)
            ppStrKakuteiDt = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【予約確定ﾕｰｻﾞCD】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrKakuUsercd</returns>ppStrKakuteiDt
    Public Property PropStrKakuUsercd()
        Get
            Return ppStrKakuUsercd
        End Get
        Set(ByVal value)
            ppStrKakuUsercd = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【予約ｽﾃｰﾀｽ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrYoyakuSts</returns>ppStrKakuUsercd
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
    ''' <returns>ppStrShisetuKbn</returns>ppStrYoyakuSts
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
    ''' <returns>ppStrStudioKbn</returns>ppStrShisetuKbn
    Public Property PropStrStudioKbn()
        Get
            Return ppStrStudioKbn
        End Get
        Set(ByVal value)
            ppStrStudioKbn = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【催事名】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrSaijiNm</returns>ppStrStudioKbn
    Public Property PropStrSaijiNm()
        Get
            Return ppStrSaijiNm
        End Get
        Set(ByVal value)
            ppStrSaijiNm = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【出演者名】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrShutsuenNm</returns>ppStrSaijiNm
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
    ''' <returns>ppStrKashiKind</returns>ppStrShutsuenNm
    Public Property PropStrKashiKind()
        Get
            Return ppStrKashiKind
        End Get
        Set(ByVal value)
            ppStrKashiKind = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【シアター利用形状】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyoType</returns>ppStrKashiKind
    Public Property PropStrRiyoType()
        Get
            Return ppStrRiyoType
        End Get
        Set(ByVal value)
            ppStrRiyoType = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【シアターワンドリンク有無】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrDrinkFlg</returns>ppStrRiyoType
    Public Property PropStrDrinkFlg()
        Get
            Return ppStrDrinkFlg
        End Get
        Set(ByVal value)
            ppStrDrinkFlg = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【シアター催事分類】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrSaijiBunrui</returns>ppStrDrinkFlg
    Public Property PropStrSaijiBunrui()
        Get
            Return ppStrSaijiBunrui
        End Get
        Set(ByVal value)
            ppStrSaijiBunrui = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【シアター定員】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrTeiin</returns>ppStrSaijiBunrui
    Public Property PropStrTeiin()
        Get
            Return ppStrTeiin
        End Get
        Set(ByVal value)
            ppStrTeiin = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者CD】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyoshaCd</returns>ppStrTeiin
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
    ''' <returns>ppStrRiyoNm</returns>ppStrRiyoshaCd
    Public Property PropStrRiyoNm()
        Get
            Return ppStrRiyoNm
        End Get
        Set(ByVal value)
            ppStrRiyoNm = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者名ｶﾅ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyoKana</returns>ppStrRiyoNm
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
    ''' <returns>ppStrSekininBusho_nm</returns>ppStrRiyoKana
    Public Property PropStrSekininBushoNm()
        Get
            Return ppStrSekininBushoNm
        End Get
        Set(ByVal value)
            ppStrSekininBushoNm = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【責任者名】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrSekininNm</returns>ppStrSekininBusho_nm
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
    ''' <returns>ppStrSekininMail</returns>ppStrSekininNm
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
    ''' <returns>ppStrDaihyoNm</returns>ppStrSekininMail
    Public Property PropStrDaihyoNm()
        Get
            Return ppStrDaihyoNm
        End Get
        Set(ByVal value)
            ppStrDaihyoNm = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者電話番号1-1】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyoTel11</returns>ppStrDaihyoNm
    Public Property PropStrRiyoTel11()
        Get
            Return ppStrRiyoTel11
        End Get
        Set(ByVal value)
            ppStrRiyoTel11 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者電話番号1-2】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyoTel12</returns>ppStrRiyoTel11
    Public Property PropStrRiyoTel12()
        Get
            Return ppStrRiyoTel12
        End Get
        Set(ByVal value)
            ppStrRiyoTel12 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者電話番号1-3】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyoTel13</returns>ppStrRiyoTel12
    Public Property PropStrRiyoTel13()
        Get
            Return ppStrRiyoTel13
        End Get
        Set(ByVal value)
            ppStrRiyoTel13 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者電話番号2-1】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyoTel21</returns>ppStrRiyoTel13
    Public Property PropStrRiyoTel21()
        Get
            Return ppStrRiyoTel21
        End Get
        Set(ByVal value)
            ppStrRiyoTel21 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者電話番号2-2】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyoTel22</returns>ppStrRiyoTel21
    Public Property PropStrRiyoTel22()
        Get
            Return ppStrRiyoTel22
        End Get
        Set(ByVal value)
            ppStrRiyoTel22 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者電話番号2-3】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyoTel23</returns>ppStrRiyoTel22
    Public Property PropStrRiyoTel23()
        Get
            Return ppStrRiyoTel23
        End Get
        Set(ByVal value)
            ppStrRiyoTel23 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者内線番号】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyoNaisen</returns>ppStrRiyoTel23
    Public Property PropStrRiyoNaisen()
        Get
            Return ppStrRiyoNaisen
        End Get
        Set(ByVal value)
            ppStrRiyoNaisen = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者FAX1-1】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyoFax11</returns>ppStrRiyoNaisen
    Public Property PropStrRiyoFax11()
        Get
            Return ppStrRiyoFax11
        End Get
        Set(ByVal value)
            ppStrRiyoFax11 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者FAX1-2】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyoFax12</returns>ppStrRiyoFax11
    Public Property PropStrRiyoFax12()
        Get
            Return ppStrRiyoFax12
        End Get
        Set(ByVal value)
            ppStrRiyoFax12 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者FAX1-3】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyoFax13</returns>ppStrRiyoFax12
    Public Property PropStrRiyoFax13()
        Get
            Return ppStrRiyoFax13
        End Get
        Set(ByVal value)
            ppStrRiyoFax13 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者郵便番号1】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyoYubin1</returns>ppStrRiyoFax13
    Public Property PropStrRiyoYubin1()
        Get
            Return ppStrRiyoYubin1
        End Get
        Set(ByVal value)
            ppStrRiyoYubin1 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者郵便番号2】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyoYubin2</returns>ppStrRiyoYubin1
    Public Property PropStrRiyoYubin2()
        Get
            Return ppStrRiyoYubin2
        End Get
        Set(ByVal value)
            ppStrRiyoYubin2 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者都道府県】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyoTodo</returns>ppStrRiyoYubin2
    Public Property PropStrRiyoTodo()
        Get
            Return ppStrRiyoTodo
        End Get
        Set(ByVal value)
            ppStrRiyoTodo = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者市区町村】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyoShiku</returns>ppStrRiyoTodo
    Public Property PropStrRiyoShiku()
        Get
            Return ppStrRiyoShiku
        End Get
        Set(ByVal value)
            ppStrRiyoShiku = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者番地】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyoBan</returns>ppStrRiyoShiku
    Public Property PropStrRiyoBan()
        Get
            Return ppStrRiyoBan
        End Get
        Set(ByVal value)
            ppStrRiyoBan = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者ビル名】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyoBuild</returns>ppStrRiyoBan
    Public Property PropStrRiyoBuild()
        Get
            Return ppStrRiyoBuild
        End Get
        Set(ByVal value)
            ppStrRiyoBuild = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者ﾚﾍﾞﾙ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyoLvl</returns>ppStrRiyoBuild
    Public Property PropStrRiyoLvl()
        Get
            Return ppStrRiyoLvl
        End Get
        Set(ByVal value)
            ppStrRiyoLvl = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【相手先コード】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrAiteCd</returns>
    Public Property PropStrAiteCd()
        Get
            Return ppStrAiteCd
        End Get
        Set(ByVal value)
            ppStrAiteCd = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【相手先名称】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrAiteNm</returns>
    Public Property PropStrAiteNm()
        Get
            Return ppStrAiteNm
        End Get
        Set(ByVal value)
            ppStrAiteNm = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【申込書送付区分】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrSendKbn</returns>
    Public Property PropStrSendKbn()
        Get
            Return ppStrSendKbn
        End Get
        Set(ByVal value)
            ppStrSendKbn = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【送付ｽﾃｰﾀｽ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrSendSts</returns>ppStrSendKbn
    Public Property PropStrSendSts()
        Get
            Return ppStrSendSts
        End Get
        Set(ByVal value)
            ppStrSendSts = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【送付日】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrSendDt</returns>ppStrSendSts
    Public Property PropStrSendDt()
        Get
            Return ppStrSendDt
        End Get
        Set(ByVal value)
            ppStrSendDt = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【返送期日】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrHensoDt</returns>ppStrSendDt
    Public Property PropStrHensoDt()
        Get
            Return ppStrHensoDt
        End Get
        Set(ByVal value)
            ppStrHensoDt = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【備考】
    ''' </summary>
    ''' <value></value>
    ''' <returns>PropBiko</returns>ppStrHensoDt
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
    ''' プロパティセット【施設利用合計金額】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrTotalRiyoKin</returns>
    ''' <remarks><para>作成情報：2015/08/26 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrTotalRiyoKin()
        Get
            Return ppStrTotalRiyoKin
        End Get
        Set(ByVal value)
            ppStrTotalRiyoKin = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用ｺﾒﾝﾄ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRiyoCom</returns>
    ''' <remarks><para>作成情報：2015/08/26 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrRiyoCom()
        Get
            Return ppStrRiyoCom
        End Get
        Set(ByVal value)
            ppStrRiyoCom = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【チケット入場料】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrTicketEnterKbn</returns>
    ''' <remarks><para>作成情報：2015/08/26 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrTicketEnterKbn()
        Get
            Return ppStrTicketEnterKbn
        End Get
        Set(ByVal value)
            ppStrTicketEnterKbn = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【チケットドリンク代】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrTicketDrinkKbn</returns>
    ''' <remarks><para>作成情報：2015/08/26 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrTicketDrinkKbn()
        Get
            Return ppStrTicketDrinkKbn
        End Get
        Set(ByVal value)
            ppStrTicketDrinkKbn = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【HP掲載】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrHpKeisai</returns>
    ''' <remarks><para>作成情報：2015/08/26 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrHpKeisai()
        Get
            Return ppStrHpKeisai
        End Get
        Set(ByVal value)
            ppStrHpKeisai = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【情報公開日】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrJohoKokaiDt</returns>
    ''' <remarks><para>作成情報：2015/08/26 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrJohoKokaiDt()
        Get
            Return ppStrJohoKokaiDt
        End Get
        Set(ByVal value)
            ppStrJohoKokaiDt = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【情報公開時間】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrJohoKokaiTime</returns>
    ''' <remarks><para>作成情報：2015/08/26 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrJohoKokaiTime()
        Get
            Return ppStrJohoKokaiTime
        End Get
        Set(ByVal value)
            ppStrJohoKokaiTime = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【公開日】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrKokaiDt</returns>
    ''' <remarks><para>作成情報：2015/08/26 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrKokaiDt()
        Get
            Return ppStrKokaiDt
        End Get
        Set(ByVal value)
            ppStrKokaiDt = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【公開時間】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrKokaiTime</returns>
    ''' <remarks><para>作成情報：2015/08/26 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrKokaiTime()
        Get
            Return ppStrKokaiTime
        End Get
        Set(ByVal value)
            ppStrKokaiTime = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【付属設備入力ｽﾃｰﾀｽ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrFinputSts</returns>
    ''' <remarks><para>作成情報：2015/08/26 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrFinputSts()
        Get
            Return ppStrFinputSts
        End Get
        Set(ByVal value)
            ppStrFinputSts = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【登録日】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrAddDt</returns>
    Public Property PropStrAddDt()
        Get
            Return ppStrAddDt
        End Get
        Set(ByVal value)
            ppStrAddDt = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【登録USER_CD】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrAddUserCd</returns>
    Public Property PropStrAddUserCd()
        Get
            Return ppStrAddUserCd
        End Get
        Set(ByVal value)
            ppStrAddUserCd = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【登録USER_NM】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrAddDt</returns>
    Public Property PropStrAddUserNm()
        Get
            Return ppStrAddUserNm
        End Get
        Set(ByVal value)
            ppStrAddUserNm = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【更新日】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrUpDt</returns>
    Public Property PropStrUpDt()
        Get
            Return ppStrUpDt
        End Get
        Set(ByVal value)
            ppStrUpDt = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【更新USER_CD】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrUpUserCd</returns>
    Public Property PropStrUpUserCd()
        Get
            Return ppStrUpUserCd
        End Get
        Set(ByVal value)
            ppStrUpUserCd = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【更新USER_NM】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrUpDt</returns>
    Public Property PropStrUpUserNm()
        Get
            Return ppStrUpUserNm
        End Get
        Set(ByVal value)
            ppStrUpUserNm = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用日リスト(DataRiyobi)】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppListRiyobi</returns>
    ''' <remarks><para>作成情報：2015/08/10 k.machida
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

    ''' <summary>
    ''' プロパティセット【利用日リスト(ArrayList String)】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppAryStrRiyoDate</returns>
    ''' <remarks><para>作成情報：2015/08/10 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropAryStrRiyoDate()
        Get
            Return ppAryStrRiyoDate
        End Get
        Set(ByVal value)
            ppAryStrRiyoDate = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【キャンセル画面から遷移】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrCancelNo</returns>
    ''' <remarks><para>作成情報：2015/08/10 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrCancelNo()
        Get
            Return ppStrCancelNo
        End Get
        Set(ByVal value)
            ppStrCancelNo = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【データセット:請求】
    ''' 格納テーブル名："BILLPAY_TBL"
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDsBillReq</returns>
    ''' <remarks><para>作成情報：2015/08/25 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDsBillReq()
        Get
            Return ppDsBillReq
        End Get
        Set(ByVal value)
            ppDsBillReq = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【HashTable:EXASプロジェクト利用料設定】
    ''' キー情報：請求依頼番号（存在しない場合は行番号）
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppHtExasPro</returns>
    ''' <remarks><para>作成情報：2015/09/04 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropHtExasPro()
        Get
            Return ppHtExasPro
        End Get
        Set(ByVal value)
            ppHtExasPro = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【HashTable:EXASプロジェクト付帯設備設定】
    ''' キー情報：請求依頼番号（存在しない場合は行番号）
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppHtExasPro</returns>
    ''' <remarks><para>作成情報：2015/09/04 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropHtExasProFutai()
        Get
            Return ppHtExasProFutai
        End Get
        Set(ByVal value)
            ppHtExasProFutai = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【DATATABLE:EXAS入金】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtExasNyukin</returns>
    ''' <remarks><para>作成情報：2015/09/02 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtExasNyukin()
        Get
            Return ppDtExasNyukin
        End Get
        Set(ByVal value)
            ppDtExasNyukin = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【データセット:承認依頼／記録】
    ''' 格納テーブル名："IRAI_RIREKI_TBL","CHECK_RIREKI_TBL"
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDsApproval</returns>
    ''' <remarks><para>作成情報：2015/08/25 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDsApproval()
        Get
            Return ppDsApproval
        End Get
        Set(ByVal value)
            ppDsApproval = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【データセット:タイムスケジュール】
    ''' 格納テーブル名："TIME_SCHEDULE_TBL"
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDsTimeSch</returns>
    ''' <remarks><para>作成情報：2015/08/25 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDsTimeSch()
        Get
            Return ppDsTimeSch
        End Get
        Set(ByVal value)
            ppDsTimeSch = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【データセット:Ticket】
    ''' 格納テーブル名："TICKET_TBL"
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDsTicket </returns>
    ''' <remarks><para>作成情報：2015/08/25 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDsTicket()
        Get
            Return ppDsTicket
        End Get
        Set(ByVal value)
            ppDsTicket = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【Hashtable:付帯】
    ''' キー情報：スラッシュ込みの日付文字列10桁
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDsFutai  </returns>
    ''' <remarks><para>作成情報：2015/08/25 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropHtFutai()
        Get
            Return ppHtFutai
        End Get
        Set(ByVal value)
            ppHtFutai = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【Hashtable:付帯】
    ''' キー情報：スラッシュ込みの日付文字列10桁
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDsFutai  </returns>
    ''' <remarks><para>作成情報：2015/08/25 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropHtFutaiDetail()
        Get
            Return ppHtFutaiDetail
        End Get
        Set(ByVal value)
            ppHtFutaiDetail = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【データセット:現金／電子】
    ''' 格納テーブル名："ALSOK_MEISAI","ALSOK_STORE"
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDsCashElectro</returns>
    ''' <remarks><para>作成情報：2015/08/25 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDsCashElectro()
        Get
            Return ppDsCashElectro
        End Get
        Set(ByVal value)
            ppDsCashElectro = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用合計】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppIntTtl</returns>
    ''' <remarks><para>作成情報：2015/08/28 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property propIntTtl()
        Get
            Return ppIntTtl
        End Get
        Set(ByVal value)
            ppIntTtl = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【消費税】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDblTax</returns>
    ''' <remarks><para>作成情報：2015/08/28 k.machida
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property propDblTax()
        Get
            Return ppDblTax
        End Get
        Set(ByVal value)
            ppDblTax = value
        End Set
    End Property

End Class
