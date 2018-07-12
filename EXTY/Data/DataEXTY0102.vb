Imports Common
Imports FarPoint.Win.Spread

''' <summary>
''' ALSOK現金入金機データ登録Dataクラス
''' </summary>
''' <remarks>ALSOK現金入金機データ登録画面で使用するのプロパティセットを行う
''' <para>作成情報：2015.08.14 h.hagiwara
''' <p>改訂情報</p>
''' </para></remarks>

Public Class DataEXTY0102

    'フォームオブジェクト
    Private ppRdoProcNew As RadioButton        ' 処理区分：新規取り込み
    Private ppRdoProcUpd As RadioButton        ' 処理区分：既存データ表示

    Private ppTxtFilePath As TextBox           ' 新規取り込み：ファイルパス表示
    Private ppBtnRef As Button                 ' 新規取り込み：ダイアログ表示ボタン
    Private ppBtnCsvInsert As Button           ' 新規取り込み：ＣＳＶ取り込み表示ボタン
    Private ppDtpDspFrom As DateTimePickerEx   ' 既存データ表示：表示期間（自）
    Private ppDtpDspTo As DateTimePickerEx     ' 既存データ表示：表示期間（至）
    Private ppBtnDataDsp As Button             ' 既存データ表示：データ表示ボタン
    Private ppBtnErrDiep As Button             ' 新規取り込み：エラー情報表示         ' 2015.12.09 ADD h.hagiwara

    Private ppVwAlsokData As FpSpread          ' ALSOK入金データSpread
    Private ppVwDegitalCashData As FpSpread    ' 電子マネー入金データSpread

    Private ppBtnAdd As Button                 ' 電子マネー入力行追加ボタン

    Private ppBtnBack As Button                ' 戻るボタン
    Private ppBtnUpdate As Button              ' 更新ボタン

    Private ppDtRejiMasta As DataTable         ' レジ情報
    Private ppDtTenpoMasta As DataTable        ' 店舗情報
    Private ppDtYoyakuData As DataTable        ' 予約情報
    Private ppDtAlsokData As DataTable         ' ALSOK入金データ取得分
    Private ppDtDegitalCashData As DataTable   ' 電子マネー入金取得分
    Private ppCbReji As CellType.ComboBoxCellType
    Private ppCbTenpo As CellType.ComboBoxCellType
    Private ppCbSaiji As CellType.ComboBoxCellType
    Private ppStrSaiji As String
    Private ppStrYoyakuNo As String            ' 予約番号                                                                              2015.12.04 ADD h.hagiwara
    Private ppStrUnionFlg As String            ' 予約情報別取得フラグ（利用日･レジ№以外から選択した予約情報取得判定 0:なし 1:あり)    2015.12.04 ADD h.hagiwara
    Private ppStrShisetuKbn As String          ' レジ施設区分                                                                          2015.12.09 ADD h.hagiwara
    Private ppStrProc As String                ' 施設区分判定                                                                          2015.12.18 ADD h.hagiwara

    ' ALSOK入金情報表 登録・更新・削除用
    Private ppRowReg As DataRow                    'データ登録／更新用：登録／更新行
    Private ppDtAlsokDataSet As DataTable          'スプレッド
    Private ppDtDegitalCashDataSet As DataTable    'スプレッド

    ' CSVファイル読み込み時の重複チェック用
    Private ppIntDupCnt As Integer
    Private ppAryDupData As ArrayList
    Private ppStrDepositCd As String
    Private ppStrSeq As String
    Private ppStrDepositDt As String
    Private ppStrRegisterCd As String
    Private ppStrTenpoCd As String
    Private ppStrRegisterNm As String
    Private ppStrTenpoNm As String
    Private ppStrErrFileNm As String           ' 最新出力エラーファイル名                   2015.12.09 ADD h.hagiwara 

    ' 共通
    Private ppDtSysDate As DateTime            ' サーバ日時
    Private ppStrUserId As String              ' 登録ユーザCD

    ''' <summary>
    ''' プロパティセット【処理区分：新規取り込み】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppRdoProcNew</returns>
    ''' <remarks><para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropRdoProcNew() As RadioButton
        Get
            Return ppRdoProcNew
        End Get
        Set(ByVal value As RadioButton)
            ppRdoProcNew = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【処理区分：既存データ表示】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppRdoProcNew</returns>
    ''' <remarks><para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropRdoProcUpd() As RadioButton
        Get
            Return ppRdoProcUpd
        End Get
        Set(ByVal value As RadioButton)
            ppRdoProcUpd = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【新規取り込み：ファイルパス表示】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppTxtFilePath</returns>
    ''' <remarks><para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropTxtFilePath() As TextBox
        Get
            Return ppTxtFilePath
        End Get
        Set(ByVal value As TextBox)
            ppTxtFilePath = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【新規取り込み：ダイアログ表示ボタン】
    ''' </summary>
    ''' <value></value>
    ''' <returns> ppBtnRef</returns>
    ''' <remarks><para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropBtnRef() As Button
        Get
            Return ppBtnRef
        End Get
        Set(ByVal value As Button)
            ppBtnRef = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【新規取り込み：ＣＳＶ取り込み表示ボタン】
    ''' </summary>
    ''' <value></value>
    ''' <returns> ppBtnCsvInsert</returns>
    ''' <remarks><para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropBtnCsvInsert() As Button
        Get
            Return ppBtnCsvInsert
        End Get
        Set(ByVal value As Button)
            ppBtnCsvInsert = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【既存データ表示：表示期間（自）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtpDspFrom</returns>
    ''' <remarks><para>作成情報：2012/06/19 kuga
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtpDspFrom() As DateTimePickerEx
        Get
            Return ppDtpDspFrom
        End Get
        Set(ByVal value As DateTimePickerEx)
            ppDtpDspFrom = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【既存データ表示：表示期間（至）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtpDspTo</returns>
    ''' <remarks><para>作成情報：2012/06/19 kuga
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtpDspTo() As DateTimePickerEx
        Get
            Return ppDtpDspTo
        End Get
        Set(ByVal value As DateTimePickerEx)
            ppDtpDspTo = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【既存データ表示：データ表示ボタン】
    ''' </summary>
    ''' <value></value>
    ''' <returns> ppBtnDataDsp</returns>
    ''' <remarks><para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropBtnDataDsp() As Button
        Get
            Return ppBtnDataDsp
        End Get
        Set(ByVal value As Button)
            ppBtnDataDsp = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【ALSOK入金データSpread】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppVwAlsokData</returns>
    ''' <remarks><para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropVwAlsokData() As FpSpread
        Get
            Return ppVwAlsokData
        End Get
        Set(ByVal value As FpSpread)
            ppVwAlsokData = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【電子マネー入金データSpread】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppVwDegitalCashData</returns>
    ''' <remarks><para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropVwDegitalCashData() As FpSpread
        Get
            Return ppVwDegitalCashData
        End Get
        Set(ByVal value As FpSpread)
            ppVwDegitalCashData = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【電子マネー入力行追加ボタン】
    ''' </summary>
    ''' <value></value>
    ''' <returns> ppBtnAdd</returns>
    ''' <remarks><para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropBtnAdd() As Button
        Get
            Return ppBtnAdd
        End Get
        Set(ByVal value As Button)
            ppBtnAdd = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【戻るボタン】
    ''' </summary>
    ''' <value></value>
    ''' <returns> ppBtnBack</returns>
    ''' <remarks><para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropBtnBack() As Button
        Get
            Return ppBtnBack
        End Get
        Set(ByVal value As Button)
            ppBtnBack = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【更新ボタン】
    ''' </summary>
    ''' <value></value>
    ''' <returns> ppBtnUpdate</returns>
    ''' <remarks><para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropBtnUpdate() As Button
        Get
            Return ppBtnUpdate
        End Get
        Set(ByVal value As Button)
            ppBtnUpdate = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【レジマスタデータ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtRejiMasta</returns>
    ''' <remarks><para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtRejiMasta() As DataTable
        Get
            Return ppDtRejiMasta
        End Get
        Set(ByVal value As DataTable)
            ppDtRejiMasta = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【店舗マスタデータ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtTenpoMasta</returns>
    ''' <remarks><para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtTenpoMasta() As DataTable
        Get
            Return ppDtTenpoMasta
        End Get
        Set(ByVal value As DataTable)
            ppDtTenpoMasta = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【ALSOK入金データ取得分】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtAlsokData</returns>
    ''' <remarks><para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtAlsokData() As DataTable
        Get
            Return ppDtAlsokData
        End Get
        Set(ByVal value As DataTable)
            ppDtAlsokData = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【電子マネー入金取得分】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtDegitalCashData</returns>
    ''' <remarks><para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtDegitalCashData() As DataTable
        Get
            Return ppDtDegitalCashData
        End Get
        Set(ByVal value As DataTable)
            ppDtDegitalCashData = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【店舗マスタデータ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtYoyakuData</returns>
    ''' <remarks><para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtYoyakuData() As DataTable
        Get
            Return ppDtYoyakuData
        End Get
        Set(ByVal value As DataTable)
            ppDtYoyakuData = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【データ登録／更新用：登録／更新行】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppRowReg</returns>
    ''' <remarks><para>作成情報：2012/07/19 r.hoshino
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropRowReg() As DataRow
        Get
            Return ppRowReg
        End Get
        Set(ByVal value As DataRow)
            ppRowReg = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【サーバ日時】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtAddDate</returns>
    ''' <remarks><para>作成情報：2015.08.14 h.hagiwara
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
    ''' プロパティセット【ログインユーザ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrUserId</returns>
    ''' <remarks><para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrUserId() As String
        Get
            Return ppStrUserId
        End Get
        Set(ByVal value As String)
            ppStrUserId = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【ＣＳＶ取り込みSpread】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtAlsokDataSet</returns>
    ''' <remarks><para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtAlsokDataSet() As DataTable
        Get
            Return ppDtAlsokDataSet
        End Get
        Set(ByVal value As DataTable)
            ppDtAlsokDataSet = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【電子マネーSpread】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtDegitalCashDataSet</returns>
    ''' <remarks><para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtDegitalCashDataSet() As DataTable
        Get
            Return ppDtDegitalCashDataSet
        End Get
        Set(ByVal value As DataTable)
            ppDtDegitalCashDataSet = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【レジ情報】
    ''' </summary>
    ''' <value></value>
    ''' <returns> ppCbReji</returns>
    ''' <remarks><para>作成情報：2015.08.14 h.hagiwara 
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropCmbReji() As CellType.ComboBoxCellType
        Get
            Return ppCbReji
        End Get
        Set(ByVal value As CellType.ComboBoxCellType)
            ppCbReji = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【店舗情報】
    ''' </summary>
    ''' <value></value>
    ''' <returns> ppCbTenpo</returns>
    ''' <remarks><para>作成情報：2015.08.14 h.hagiwara 
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropCmbTenpo() As CellType.ComboBoxCellType
        Get
            Return ppCbTenpo
        End Get
        Set(ByVal value As CellType.ComboBoxCellType)
            ppCbTenpo = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【予約情報コンボ】
    ''' </summary>
    ''' <value></value>
    ''' <returns> ppCbSaiji</returns>
    ''' <remarks><para>作成情報：2015.08.14 h.hagiwara 
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropCmbSaiji() As CellType.ComboBoxCellType
        Get
            Return ppCbSaiji
        End Get
        Set(ByVal value As CellType.ComboBoxCellType)
            ppCbSaiji = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【重複件数】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppIntDupCnt</returns>
    ''' <remarks><para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropIntDupCnt() As Integer
        Get
            Return ppIntDupCnt
        End Get
        Set(ByVal value As Integer)
            ppIntDupCnt = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【重複件数】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppAryDupData</returns>
    ''' <remarks><para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropAryDupData() As ArrayList
        Get
            Return ppAryDupData
        End Get
        Set(ByVal value As ArrayList)
            ppAryDupData = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【重複チェック：入金機コード】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrDepositCd</returns>
    ''' <remarks><para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrDepositCd() As String
        Get
            Return ppStrDepositCd
        End Get
        Set(ByVal value As String)
            ppStrDepositCd = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【重複チェック：連番】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrSeq</returns>
    ''' <remarks><para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrSeq() As String
        Get
            Return ppStrSeq
        End Get
        Set(ByVal value As String)
            ppStrSeq = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【予約情報取得用：利用日】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrDepositDt</returns>
    ''' <remarks><para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrDepositDt() As String
        Get
            Return ppStrDepositDt
        End Get
        Set(ByVal value As String)
            ppStrDepositDt = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【予約情報取得用：レジコード】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRegisterCd</returns>
    ''' <remarks><para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrRegisterCd() As String
        Get
            Return ppStrRegisterCd
        End Get
        Set(ByVal value As String)
            ppStrRegisterCd = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【予約情報取得用：店舗コード】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrTenpoCd</returns>
    ''' <remarks><para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrTenpoCd() As String
        Get
            Return ppStrTenpoCd
        End Get
        Set(ByVal value As String)
            ppStrTenpoCd = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【予約情報取得用：レジ名】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrRegisterNm</returns>
    ''' <remarks><para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrRegisterNm() As String
        Get
            Return ppStrRegisterNm
        End Get
        Set(ByVal value As String)
            ppStrRegisterNm = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【予約情報取得用：店舗名】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrTenpoNm</returns>
    ''' <remarks><para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrTenpoNm() As String
        Get
            Return ppStrTenpoNm
        End Get
        Set(ByVal value As String)
            ppStrTenpoNm = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【予約情報取得用：催事名】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrSaiji</returns>
    ''' <remarks><para>作成情報：2015.08.14 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrSaiji() As String
        Get
            Return ppStrSaiji
        End Get
        Set(ByVal value As String)
            ppStrSaiji = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【予約情報取得用：予約番号】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrYoyakuNo</returns>
    ''' <remarks><para>作成情報：2015.12.04 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrYoyakuNo() As String
        Get
            Return ppStrYoyakuNo
        End Get
        Set(ByVal value As String)
            ppStrYoyakuNo = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【予約情報取得用：予約情報取得フラグ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrUnionFlg</returns>
    ''' <remarks><para>作成情報：2015.12.04 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrUnionFlg() As String
        Get
            Return ppStrUnionFlg
        End Get
        Set(ByVal value As String)
            ppStrUnionFlg = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【新規取り込み：エラー内容表示ボタン】
    ''' </summary>
    ''' <value></value>
    ''' <returns> ppBtnErrDiep</returns>
    ''' <remarks><para>作成情報：2015.12.09 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropBtnErrDiep() As Button
        Get
            Return ppBtnErrDiep
        End Get
        Set(ByVal value As Button)
            ppBtnErrDiep = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【新規取り込み：最新出力エラーファイル名】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrErrFileNm</returns>
    ''' <remarks><para>作成情報：2015.12.09 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrErrFileNm() As String
        Get
            Return ppStrErrFileNm
        End Get
        Set(ByVal value As String)
            ppStrErrFileNm = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【レジ取得：施設区分】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrShisetuKbn</returns>
    ''' <remarks><para>作成情報：2015.12.09 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrShisetuKbn() As String
        Get
            Return ppStrShisetuKbn
        End Get
        Set(ByVal value As String)
            ppStrShisetuKbn = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【予約情報取得：施設区分】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrProc</returns>
    ''' <remarks><para>作成情報：2015.12.18 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrProc() As String
        Get
            Return ppStrProc
        End Get
        Set(ByVal value As String)
            ppStrProc = value
        End Set
    End Property

End Class
