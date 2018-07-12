Imports FarPoint.Win.Spread
Public Class DataEXTB0103

    '帳票印刷用
    Private ppStrReserveNo_Output As String             '承認番号(予約番号)
    Private ppStrUserName_Output As String              '会社・団体名(利用者名)
    Private ppStrDelegateName_Output As String          '代表者名
    Private ppStrChargeName_Output As String            '利用責任者(責任者名)
    Private ppStrEventInfo_Output As String             '利用内容(催事分類)
    Private ppStrEventName_Output As String             '催事名
    Private ppStrUseType_Output As String               '利用形状
    Private ppStrCapacity_Output As String              '定員
    Private ppStrDrinkFee_Output As String              'ドリンク代徴収(ドリンク代)
    Private ppStrUseDays_Output As String               '利用日時
    Private ppStrCountUseDays_Output As String          '利用日数
    Private ppLngTheaterFee_Output As Long              '利用料金
    Private ppLngAdjustFee_Output As Long               '調整額
    Private ppStrCalculateDay_Output As String          '税率計算日
    Private ppStrTax_Output As String                   '消費税
    Private ppStrOutputDay_Output As String             '出力日    
    Private ppDtUseApproval_Output As DataTable         '利用承認書
    Private ppDtUseDetails_Output As DataTable          '利用明細
    Private ppDtUseDetailsNoTax_Output As DataTable     '利用明細(税込分)
    Private ppStrRentalClass_Output As String           '貸出種別

    '画面情報
    Private ppLblReserveNo As Label                     '予約番号
    Private ppGrpEventInfo As GroupBox                  '催事情報グループボックス
    Private ppTxtEventName As TextBox                   '催事名
    Private ppTxtPerformerName As TextBox               '出演者名
    Private ppPnlUseType As Panel                       '利用形状パネル
    Private ppRdoSeating As RadioButton                 '利用形状(イベント（着席）)
    Private ppRdoStanding As RadioButton                '利用形状(イベント（ｽﾀﾝﾃﾞｨﾝｸﾞ）)
    Private ppRdoAnomaly As RadioButton                 '利用形状(イベント（変則）)
    Private ppRdoEvent As RadioButton                   '利用形状(催事)
    Private ppTxtCapacity As TextBox                    '利用形状(定員)
    Private ppPnlEventClass As Panel                    '催事分類パネル
    Private ppRdoMusic As RadioButton                   '催事分類(音楽)
    Private ppRdoTheater As RadioButton                 '催事分類(演劇)
    Private ppRdoEntertainment As RadioButton           '催事分類(演芸)
    Private ppRdoBusiness As RadioButton                '催事分類(ビジネス)
    Private ppTxtUserName As TextBox                    '利用者名（会社名・団体名）
    Private ppTxtDelegateName As TextBox                '代表者名
    Private ppTxtChargeName As TextBox                  '責任者名
    Private ppVwUseDays As FpSpread                     '利用日時スプレッドシート
    Private ppVwClaimPayment As FpSpread                '請求／入金スプレッドシート
    Private ppPnlDrinkFee As Panel                      'ドリンク代パネル
    Private ppRdoDrinkIn As RadioButton                 'ドリンク代(込)
    Private ppRdoDrinkOut As RadioButton                'ドリンク代(別)
    Private ppRdoNothing As RadioButton                 'ドリンク代(無し)
    Private ppVwPrintSheet As FpSpread                  '印刷用スプレッドシート
    Private ppVwIncident As FpSpread                    '付帯設備スプレッドシート
    Private ppPnlRentalClass As Panel                   '貸出種別パネル

    Private ppBlnNoTaxFlg As Boolean                    '税無フラグ
    Private ppBlnSumFlg As Boolean                      '合算フラグ
    Private ppStrClickedIncident As String              '付帯設備スプレッドシートにてボタンがクリックされた行の日付
    Private ppStrClickedRiyobi As String                ' 利用日時でクリックされた行の日付         ' 2015.12.16 ADD h.hagiwara
    '請求データＣＳＶ作成用
    Private ppStrGrpKey As String                       ' 集計キー
    Private ppStrNoTaxFlg As String                     ' 税無フラグ
    Private ppStrAitecd As String                       ' 相手先コード
    Private ppStrPostno As String                       ' 郵便番号
    Private ppStrAddr1 As String                        ' 住所１
    Private ppStrAddr2 As String                        ' 住所２
    Private ppStrAitenm As String                       ' 相手先名

    ''' <summary>
    ''' プロパティセット【帳票：承認番号(予約番号)】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrReserveNo_Output</returns>
    ''' <remarks><para>作成情報：2015/08/10 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrReserveNo_Output() As String
        Get
            Return ppStrReserveNo_Output
        End Get
        Set(ByVal value As String)
            ppStrReserveNo_Output = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【帳票：会社・団体名(利用者名)】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrUserName_Output</returns>
    ''' <remarks><para>作成情報：2015/08/10 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrUserName_Output() As String
        Get
            Return ppStrUserName_Output
        End Get
        Set(ByVal value As String)
            ppStrUserName_Output = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【帳票：代表者名】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrDelegateName_Output</returns>
    ''' <remarks><para>作成情報：2015/08/10 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrDelegateName_Output() As String
        Get
            Return ppStrDelegateName_Output
        End Get
        Set(ByVal value As String)
            ppStrDelegateName_Output = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【帳票：利用責任者(責任者名)】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrChargeName_Output</returns>
    ''' <remarks><para>作成情報：2015/08/10 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrChargeName_Output() As String
        Get
            Return ppStrChargeName_Output
        End Get
        Set(ByVal value As String)
            ppStrChargeName_Output = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【帳票：利用内容(催事分類)】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrEventInfo_Output</returns>
    ''' <remarks><para>作成情報：2015/08/10 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrEventInfo_Output() As String
        Get
            Return ppStrEventInfo_Output
        End Get
        Set(ByVal value As String)
            ppStrEventInfo_Output = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【帳票：催事名】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrEventName_Output</returns>
    ''' <remarks><para>作成情報：2015/08/10 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrEventName_Output() As String
        Get
            Return ppStrEventName_Output
        End Get
        Set(ByVal value As String)
            ppStrEventName_Output = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【帳票：利用形状】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrUseType_Output</returns>
    ''' <remarks><para>作成情報：2015/08/10 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrUseType_Output() As String
        Get
            Return ppStrUseType_Output
        End Get
        Set(ByVal value As String)
            ppStrUseType_Output = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【帳票：定員】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrCapacity_Output</returns>
    ''' <remarks><para>作成情報：2015/08/10 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrCapacity_Output() As String
        Get
            Return ppStrCapacity_Output
        End Get
        Set(ByVal value As String)
            ppStrCapacity_Output = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【帳票：ドリンク代徴収(ドリンク代)】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrDrinkFee_Output</returns>
    ''' <remarks><para>作成情報：2015/08/10 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrDrinkFee_Output() As String
        Get
            Return ppStrDrinkFee_Output
        End Get
        Set(ByVal value As String)
            ppStrDrinkFee_Output = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【帳票：利用日時】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrUseDays_Output</returns>
    ''' <remarks><para>作成情報：2015/08/10 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrUseDays_Output() As String
        Get
            Return ppStrUseDays_Output
        End Get
        Set(ByVal value As String)
            ppStrUseDays_Output = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【帳票：利用日数】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrCountUseDays_Output</returns>
    ''' <remarks><para>作成情報：2015/08/10 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrCountUseDays_Output() As String
        Get
            Return ppStrCountUseDays_Output
        End Get
        Set(ByVal value As String)
            ppStrCountUseDays_Output = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【帳票：利用料金】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppLngTheaterFee_Output</returns>
    ''' <remarks><para>作成情報：2015/08/17 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropLngTheaterFee_Output() As Long
        Get
            Return ppLngTheaterFee_Output
        End Get
        Set(ByVal value As Long)
            ppLngTheaterFee_Output = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【調整額】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppLngAdjustFee_Output</returns>
    ''' <remarks><para>作成情報：2015/08/18 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropLngAdjustFee_Output() As Long
        Get
            Return ppLngAdjustFee_Output
        End Get
        Set(ByVal value As Long)
            ppLngAdjustFee_Output = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【帳票：税率計算日】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrCalculateDay_Output</returns>
    ''' <remarks><para>作成情報：2015/08/14 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrCalculateDay_Output() As String
        Get
            Return ppStrCalculateDay_Output
        End Get
        Set(ByVal value As String)
            ppStrCalculateDay_Output = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【帳票：消費税】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrTax_Output</returns>
    ''' <remarks><para>作成情報：2015/08/10 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrTax_Output() As String
        Get
            Return ppStrTax_Output
        End Get
        Set(ByVal value As String)
            ppStrTax_Output = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【帳票：出力日】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrOutputDay_Output</returns>
    ''' <remarks><para>作成情報：2015/08/14 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrOutputDay_Output() As String
        Get
            Return ppStrOutputDay_Output
        End Get
        Set(ByVal value As String)
            ppStrOutputDay_Output = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【帳票：利用承認書データテーブル】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtUseDetails_Output</returns>
    ''' <remarks><para>作成情報：2015/08/28 y.naganuma
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtUseApproval_Output() As DataTable
        Get
            Return ppDtUseApproval_Output
        End Get
        Set(ByVal value As DataTable)
            ppDtUseApproval_Output = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【帳票：利用明細データテーブル】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtUseDetails_Output</returns>
    ''' <remarks><para>作成情報：2015/08/14 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtUseDetails_Output() As DataTable
        Get
            Return ppDtUseDetails_Output
        End Get
        Set(ByVal value As DataTable)
            ppDtUseDetails_Output = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【帳票：利用明細（税込分）データテーブル】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtUseDetailsNoTax_Output</returns>
    ''' <remarks><para>作成情報：2015/08/14 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtUseDetailsNoTax_Output() As DataTable
        Get
            Return ppDtUseDetailsNoTax_Output
        End Get
        Set(ByVal value As DataTable)
            ppDtUseDetailsNoTax_Output = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【帳票：貸出種別（判定）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRentalClass_Output</returns>
    ''' <remarks><para>作成情報：2015/08/17 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrRentalClass_Output() As String
        Get
            Return ppStrRentalClass_Output
        End Get
        Set(ByVal value As String)
            ppStrRentalClass_Output = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【予約番号】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppLblReserveNo</returns>
    ''' <remarks><para>作成情報：2015/08/10 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropLblReserveNo() As Label
        Get
            Return ppLblReserveNo
        End Get
        Set(ByVal value As Label)
            ppLblReserveNo = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【催事情報グループボックス】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppGrpEventInfo</returns>
    ''' <remarks><para>作成情報：2015/08/10 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropGrpEventInfo() As GroupBox
        Get
            Return ppGrpEventInfo
        End Get
        Set(ByVal value As GroupBox)
            ppGrpEventInfo = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【催事名】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppTxtEventName</returns>
    ''' <remarks><para>作成情報：2015/08/10 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropTxtEventName() As TextBox
        Get
            Return ppTxtEventName
        End Get
        Set(ByVal value As TextBox)
            ppTxtEventName = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【出演者名】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppTxtPerformerName</returns>
    ''' <remarks><para>作成情報：2015/08/10 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropTxtPerformerName() As TextBox
        Get
            Return ppTxtPerformerName
        End Get
        Set(ByVal value As TextBox)
            ppTxtPerformerName = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用形状パネル】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppPnlUseType</returns>
    ''' <remarks><para>作成情報：2015/08/10 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropPnlUseType() As Panel
        Get
            Return ppPnlUseType
        End Get
        Set(ByVal value As Panel)
            ppPnlUseType = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用形状(イベント（着席）)】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppRdoSeating</returns>
    ''' <remarks><para>作成情報：2015/08/10 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropRdoSeating() As RadioButton
        Get
            Return ppRdoSeating
        End Get
        Set(ByVal value As RadioButton)
            ppRdoSeating = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用形状(イベント（ｽﾀﾝﾃﾞｨﾝｸﾞ）)】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppRdoStanding</returns>
    ''' <remarks><para>作成情報：2015/08/10 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropRdoStanding() As RadioButton
        Get
            Return ppRdoStanding
        End Get
        Set(ByVal value As RadioButton)
            ppRdoStanding = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用形状(イベント（変則）)】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppRdoAnomaly</returns>
    ''' <remarks><para>作成情報：2015/08/10 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropRdoAnomaly() As RadioButton
        Get
            Return ppRdoAnomaly
        End Get
        Set(ByVal value As RadioButton)
            ppRdoAnomaly = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用形状(催事)】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppRdoEvent</returns>
    ''' <remarks><para>作成情報：2015/08/10 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropRdoEvent() As RadioButton
        Get
            Return ppRdoEvent
        End Get
        Set(ByVal value As RadioButton)
            ppRdoEvent = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用形状(定員)】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppTxtCapacity</returns>
    ''' <remarks><para>作成情報：2015/08/10 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropTxtCapacity() As TextBox
        Get
            Return ppTxtCapacity
        End Get
        Set(ByVal value As TextBox)
            ppTxtCapacity = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【催事分類パネル】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppPnlEventClass</returns>
    ''' <remarks><para>作成情報：2015/08/10 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropPnlEventClass() As Panel
        Get
            Return ppPnlEventClass
        End Get
        Set(ByVal value As Panel)
            ppPnlEventClass = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【催事分類(音楽)】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppRdoMusic</returns>
    ''' <remarks><para>作成情報：2015/08/10 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropRdoMusic() As RadioButton
        Get
            Return ppRdoMusic
        End Get
        Set(ByVal value As RadioButton)
            ppRdoMusic = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【催事分類(演劇)】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppRdoTheater</returns>
    ''' <remarks><para>作成情報：2015/08/10 d.sonoda
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
    ''' プロパティセット【催事分類(演芸)】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppRdoEntertainment</returns>
    ''' <remarks><para>作成情報：2015/08/10 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropRdoEntertainment() As RadioButton
        Get
            Return ppRdoEntertainment
        End Get
        Set(ByVal value As RadioButton)
            ppRdoEntertainment = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【催事分類(ビジネス)】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppRdoBusiness</returns>
    ''' <remarks><para>作成情報：2015/08/10 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropRdoBusiness() As RadioButton
        Get
            Return ppRdoBusiness
        End Get
        Set(ByVal value As RadioButton)
            ppRdoBusiness = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用者名（会社名・団体名）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppTxtUserName</returns>
    ''' <remarks><para>作成情報：2015/08/10 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropTxtUserName() As TextBox
        Get
            Return ppTxtUserName
        End Get
        Set(ByVal value As TextBox)
            ppTxtUserName = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【代表者名】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppTxtDelegateName</returns>
    ''' <remarks><para>作成情報：2015/08/10 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropTxtDelegateName() As TextBox
        Get
            Return ppTxtDelegateName
        End Get
        Set(ByVal value As TextBox)
            ppTxtDelegateName = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【責任者名】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppTxtChargeName</returns>
    ''' <remarks><para>作成情報：2015/08/10 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropTxtChargeName() As TextBox
        Get
            Return ppTxtChargeName
        End Get
        Set(ByVal value As TextBox)
            ppTxtChargeName = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用日時スプレッドシート】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppVwUseDays</returns>
    ''' <remarks><para>作成情報：2015/08/10 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropVwUseDays() As FpSpread
        Get
            Return ppVwUseDays
        End Get
        Set(ByVal value As FpSpread)
            ppVwUseDays = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【請求／入金スプレッドシート】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppVwClaimPayment</returns>
    ''' <remarks><para>作成情報：2015/08/10 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropVwClaimPayment() As FpSpread
        Get
            Return ppVwClaimPayment
        End Get
        Set(ByVal value As FpSpread)
            ppVwClaimPayment = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【ドリンク代パネル】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppPnlDrinkFee</returns>
    ''' <remarks><para>作成情報：2015/08/10 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropPnlDrinkFee() As Panel
        Get
            Return ppPnlDrinkFee
        End Get
        Set(ByVal value As Panel)
            ppPnlDrinkFee = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【ドリンク代(込)】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppRdoDrinkIn</returns>
    ''' <remarks><para>作成情報：2015/08/10 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropRdoDrinkIn() As RadioButton
        Get
            Return ppRdoDrinkIn
        End Get
        Set(ByVal value As RadioButton)
            ppRdoDrinkIn = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【ドリンク代(別)】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppRdoDrinkOut</returns>
    ''' <remarks><para>作成情報：2015/08/10 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropRdoDrinkOut() As RadioButton
        Get
            Return ppRdoDrinkOut
        End Get
        Set(ByVal value As RadioButton)
            ppRdoDrinkOut = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【ドリンク代(無し)】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppRdoNothing</returns>
    ''' <remarks><para>作成情報：2015/08/10 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropRdoNothing() As RadioButton
        Get
            Return ppRdoNothing
        End Get
        Set(ByVal value As RadioButton)
            ppRdoNothing = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【印刷用スプレッドシート】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppVwPrintSheet</returns>
    ''' <remarks><para>作成情報：2015/08/10 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropVwPrintSheet() As FpSpread
        Get
            Return ppVwPrintSheet
        End Get
        Set(ByVal value As FpSpread)
            ppVwPrintSheet = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【付帯設備スプレッドシート】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppVwIncident</returns>
    ''' <remarks><para>作成情報：2015/08/14 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropVwIncident() As FpSpread
        Get
            Return ppVwIncident
        End Get
        Set(ByVal value As FpSpread)
            ppVwIncident = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【貸出種別パネル】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppPnlRentalClass</returns>
    ''' <remarks><para>作成情報：2015/08/17 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropPnlRentalClass() As Panel
        Get
            Return ppPnlRentalClass
        End Get
        Set(ByVal value As Panel)
            ppPnlRentalClass = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【税無フラグ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppBlnNoTaxFlg</returns>
    ''' <remarks><para>作成情報：2015/08/14 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropBlnNoTaxFlg() As Boolean
        Get
            Return ppBlnNoTaxFlg
        End Get
        Set(ByVal value As Boolean)
            ppBlnNoTaxFlg = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【合算フラグ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppBlnSumFlg</returns>
    ''' <remarks><para>作成情報：2015/08/17 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropBlnSumFlg() As Boolean
        Get
            Return ppBlnSumFlg
        End Get
        Set(ByVal value As Boolean)
            ppBlnSumFlg = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【付帯設備スプレッドシートにてボタンがクリックされた行の日付】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrClickedIncident</returns>
    ''' <remarks><para>作成情報：2015/08/18 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrClickedIncident() As String
        Get
            Return ppStrClickedIncident
        End Get
        Set(ByVal value As String)
            ppStrClickedIncident = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【EXAS請求情報作成時の消費税率判定用集計キー】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrGrpKey</returns>
    ''' <remarks><para>作成情報：2015/11/11 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrGrpKey() As String
        Get
            Return ppStrGrpKey
        End Get
        Set(ByVal value As String)
            ppStrGrpKey = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【EXAS請求情報作成時の消費税率判定用税抜フラグ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrNoTaxFlg</returns>
    ''' <remarks><para>作成情報：2015/11/11 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrNoTaxFlg() As String
        Get
            Return ppStrNoTaxFlg
        End Get
        Set(ByVal value As String)
            ppStrNoTaxFlg = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【EXAS請求情報用：相手先コード】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrAitecd</returns>
    ''' <remarks><para>作成情報：2015/11/11 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrAitecd() As String
        Get
            Return ppStrAitecd
        End Get
        Set(ByVal value As String)
            ppStrAitecd = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【EXAS請求情報用：郵便番号】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrPostno</returns>
    ''' <remarks><para>作成情報：2015/11/11 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrPostno() As String
        Get
            Return ppStrPostno
        End Get
        Set(ByVal value As String)
            ppStrPostno = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【EXAS請求情報用：住所１】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrAddr1</returns>
    ''' <remarks><para>作成情報：2015/11/11 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrAddr1() As String
        Get
            Return ppStrAddr1
        End Get
        Set(ByVal value As String)
            ppStrAddr1 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【EXAS請求情報用：住所２】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrAddr2</returns>
    ''' <remarks><para>作成情報：2015/11/11 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrAddr2() As String
        Get
            Return ppStrAddr2
        End Get
        Set(ByVal value As String)
            ppStrAddr2 = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【EXAS請求情報用：相手先名】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrAitenm</returns>
    ''' <remarks><para>作成情報：2015/11/11 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrAitenm() As String
        Get
            Return ppStrAitenm
        End Get
        Set(ByVal value As String)
            ppStrAitenm = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用日時でクリックされた行の日付】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrClickedRiyobi</returns>
    ''' <remarks><para>作成情報：2015.12.16.h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrClickedRiyobi() As String
        Get
            Return ppStrClickedRiyobi
        End Get
        Set(ByVal value As String)
            ppStrClickedRiyobi = value
        End Set
    End Property

End Class
