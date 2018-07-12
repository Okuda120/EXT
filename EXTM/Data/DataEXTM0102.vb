Imports FarPoint.Win.Spread

Public Class DataEXTM0102
    '初期化フラグ
    Private ppInitFlg As Boolean
    'コンボボックス初期作成用フラグ
    Private ppInitCmbFlg As Boolean

    '編集内容をもとに新規登録ボタンフラグ
    Private ppNewEntryBtnFlg As Boolean

    ' 期間FromTo
    Private ppYearFrom As TextBox  'from年
    Private ppMonthFrom As TextBox 'from月
    Private ppYearTo As TextBox    'To年
    Private ppMonthTo As TextBox   'To月
    ' 期間FRoMTo（更新用退避)
    Private ppFromYearUp As String     ' From年
    Private ppFromMonthUp As String    ' From月
    Private ppToYearUp As String       ' To年
    Private ppToMonthUp As String      ' To月
    Private ppMaxBunruiCd As String
    Private ppMaxFutaiBunruiCd As Integer
    Private ppUpdBunruiCd As String
    Private ppBunruiCd As String
    Private ppRow As Integer

    ' 2015.12.25 ADD START↓ h.hagiwara 利用料用コピー元設定
    Private ppstrKikan_From As String
    Private ppstrKikan_To As String
    Private ppstrShisetu_kbn As String
    Private ppstrBunrui_Cd As String
    ' 2015.12.25 ADD END↑ h.hagiwara 利用料用コピー元設定

    ' 2016.06.24 ADD START↓ h.hagiwara コンボリスト設定方法変更対応
    Private ppstrSelKanjo As String              ' データテーブルからの検索用勘定科目
    Private ppstrSelSaimoku As String            ' データテーブルからの検索用細目
    Private ppstrSelUchiwake As String           ' データテーブルからの検索用内訳
    ' 2016.06.24 ADD END↑ h.hagiwara コンボリスト設定方法変更対応

    'コンボボックス
    Private ppFinishedFromTo As ComboBox ' 編集済みの料金設定FromToコンボボックス

    'ラジオボタン
    Private ppNewBtn As RadioButton      '新規に料金を設定する
    Private ppFinishedBtn As RadioButton '設定済みの料金を編集する
    Private ppTheaterBtn As RadioButton  'シアター
    Private ppStudioBtn As RadioButton   'スタジオ

    'ボタン
    Private ppBackBtn As Button         '戻るボタン
    Private ppNewEntryBtn As Button     '登録内容をもとに新規登録
    Private ppEntryBtn As Button        '登録ボタン

    'テーブル
    Private ppDtKamokuMst As DataTable    '科目マスタ
    Private ppDtFbunruiMst As DataTable   '付帯設備分類マスタ
    Private ppDtFutaiMst As DataTable     '付帯設備マスタ
    Private ppDtFinishedFromTo As DataTable  '編集済みの対象期間
    Private ppDtKanzyoCd As DataTable       '科目コードコンボボックス用
    ' 2016.04.28 DEL START↓ h.hagiwara レスポンス改善
    'Private ppDtKariKanzyoCd As DataTable   '借方科目コードコンボボックス用
    ' 2016.04.28 DEL END↑ h.hagiwara レスポンス改善
    Private ppDtFutaiMstDsp As DataTable    '付帯設備マスタ（表示設定用） 20151023 
    Private ppDtFutaiMstEscp As DataTable   '付帯設備マスタ（退避用）     20151023
    Private ppDtCopyKamokuMst As DataTable    ' 利用料用科目マスタ      ' 2015.12.25 ADD h.hagiwara
    Private ppDtSaimoku As DataTable          ' 科目マスタ(細目)        ' 2016.06.24 ADD h.hagiwara コンボリスト設定方法変更対応        
    Private ppDtUchiwake As DataTable         ' 科目マスタ(内訳)        ' 2016.06.24 ADD h.hagiwara コンボリスト設定方法変更対応        
    Private ppDtShosai As DataTable           ' 科目マスタ(詳細)        ' 2016.06.24 ADD h.hagiwara コンボリスト設定方法変更対応        

    'シート
    Private ppVwGroupingSheet As FpSpread   '分類表
    Private ppVwFutaiSheet As FpSpread      '付帯設備表


    ''' <summary>
    ''' プロパティセット【ppInitCmbFlg】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppInitCmbFlg</returns>
    ''' <remarks><para>作成情報：2015/08/21 yu.satoh
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropInitCmbFlg() As Boolean
        Get
            Return ppInitCmbFlg
        End Get
        Set(ByVal value As Boolean)
            ppInitCmbFlg = value
        End Set
    End Property


    ''' <summary>
    ''' プロパティセット【ppNewEntryBtnFlg】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppNewEntryBtnFlg</returns>
    ''' <remarks><para>作成情報：2015/08/20 yu.satoh
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropNewEntryBtnFlg() As Boolean
        Get
            Return ppNewEntryBtnFlg
        End Get
        Set(ByVal value As Boolean)
            ppNewEntryBtnFlg = value
        End Set
    End Property

    ' 2016.04.28 DEL START↓ h.hagiwara レスポンス改善
    ' ''' <summary>
    ' ''' プロパティセット【ppDtKariKanzyoCd】
    ' ''' </summary>
    ' ''' <value></value>
    ' ''' <returns>ppDtKariKanzyoCd</returns>
    ' ''' <remarks><para>作成情報：2015/08/20 yu.satoh
    ' ''' <p>改訂情報:</p>
    ' ''' </para></remarks>
    'Public Property PropDtKariKanzyoCd() As DataTable
    '    Get
    '        Return ppDtKariKanzyoCd
    '    End Get
    '    Set(ByVal value As DataTable)
    '        ppDtKariKanzyoCd = value
    '    End Set
    'End Property
    ' 2016.04.28 DEL END↑ h.hagiwara レスポンス改善

    ''' <summary>
    ''' プロパティセット【ppDtKanzyoCd】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtKanzyoCd</returns>
    ''' <remarks><para>作成情報：2015/08/20 yu.satoh
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtKanzyoCd() As DataTable
        Get
            Return ppDtKanzyoCd
        End Get
        Set(ByVal value As DataTable)
            ppDtKanzyoCd = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【ppInitFlg】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppInitFlg</returns>
    ''' <remarks><para>作成情報：2015/08/17 yu.satoh
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropInitFlg() As Boolean
        Get
            Return ppInitFlg
        End Get
        Set(ByVal value As Boolean)
            ppInitFlg = value
        End Set
    End Property


    ''' <summary>
    ''' プロパティセット【ppDtFinishedFromTo】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppInitFlg</returns>
    ''' <remarks><para>作成情報：2015/08/17 yu.satoh
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtFinishedFromTo() As DataTable
        Get
            Return ppDtFinishedFromTo
        End Get
        Set(ByVal value As DataTable)
            ppDtFinishedFromTo = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【ppYearFrom】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppYearFrom</returns>
    ''' <remarks><para>作成情報：2015/08/11  yu.satoh
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropYearFrom() As TextBox
        Get
            Return ppYearFrom
        End Get
        Set(ByVal value As TextBox)
            ppYearFrom = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【ppMonthFrom】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppMonthFrom</returns>
    ''' <remarks><para>作成情報：2015/08/11  yu.satoh
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropMonthFrom() As TextBox
        Get
            Return ppMonthFrom
        End Get
        Set(ByVal value As TextBox)
            ppMonthFrom = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【ppYearTo】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppYearTo</returns>
    ''' <remarks><para>作成情報：2015/08/11  yu.satoh
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropYearTo() As TextBox
        Get
            Return ppYearTo
        End Get
        Set(ByVal value As TextBox)
            ppYearTo = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【ppMonthTo】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppMonthTo</returns>
    ''' <remarks><para>作成情報：2015/08/11  yu.satoh
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropMonthTo() As TextBox
        Get
            Return ppMonthTo
        End Get
        Set(ByVal value As TextBox)
            ppMonthTo = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【ppFinishedFromTo】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppFinishedFromTo</returns>
    ''' <remarks><para>作成情報：2015/08/11  yu.satoh
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropFinishedFromTo() As ComboBox
        Get
            Return ppFinishedFromTo
        End Get
        Set(ByVal value As ComboBox)
            ppFinishedFromTo = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【ppNewBtn】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppNewBtn</returns>
    ''' <remarks><para>作成情報：2015/08/11  yu.satoh
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropNewBtn() As RadioButton
        Get
            Return ppNewBtn
        End Get
        Set(ByVal value As RadioButton)
            ppNewBtn = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【ppFinishedBtn】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppFinishedBtn</returns>
    ''' <remarks><para>作成情報：2015/08/11  yu.satoh
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropFinishedBtn() As RadioButton
        Get
            Return ppFinishedBtn
        End Get
        Set(ByVal value As RadioButton)
            ppFinishedBtn = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【ppTheaterBtn】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppTheaterBtn</returns>
    ''' <remarks><para>作成情報：2015/08/11  yu.satoh
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropTheaterBtn() As RadioButton
        Get
            Return ppTheaterBtn
        End Get
        Set(ByVal value As RadioButton)
            ppTheaterBtn = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【ppStudioBtn】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStudioBtn</returns>
    ''' <remarks><para>作成情報：2015/08/11  yu.satoh
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStudioBtn() As RadioButton
        Get
            Return ppStudioBtn
        End Get
        Set(ByVal value As RadioButton)
            ppStudioBtn = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【EXTM科目マスタ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtKamokuMst</returns>
    ''' <remarks><para>作成情報：2015/08/11  yu.satoh
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtKamokuMst() As DataTable
        Get
            Return ppDtKamokuMst
        End Get
        Set(ByVal value As DataTable)
            ppDtKamokuMst = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【EXTM付帯設備分類】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtFbunruiMst</returns>
    ''' <remarks><para>作成情報：2015/08/11  yu.satoh
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtFbunruiMst() As DataTable
        Get
            Return ppDtFbunruiMst
        End Get
        Set(ByVal value As DataTable)
            ppDtFbunruiMst = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【EXTM付帯設備分類】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtFutaiMst</returns>
    ''' <remarks><para>作成情報：2015/08/14  yu.satoh
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtFutaiMst() As DataTable
        Get
            Return ppDtFutaiMst
        End Get
        Set(ByVal value As DataTable)
            ppDtFutaiMst = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【ppVwGroupingSheet】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppVwGroupingSheet</returns>
    ''' <remarks><para>作成情報：2015/08/11  yu.satoh
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropVwGroupingSheet() As FpSpread
        Get
            Return ppVwGroupingSheet
        End Get
        Set(ByVal value As FpSpread)
            ppVwGroupingSheet = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【ppVwFutaiSheet】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppVwFutaiSheet</returns>
    ''' <remarks><para>作成情報：2015/08/11  yu.satoh
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropVwFutaiSheet() As FpSpread
        Get
            Return ppVwFutaiSheet
        End Get
        Set(ByVal value As FpSpread)
            ppVwFutaiSheet = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【ppBackBtn】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppBackBtn</returns>
    ''' <remarks><para>作成情報：2015/08/11  yu.satoh
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropBackBtn() As Button
        Get
            Return ppBackBtn
        End Get
        Set(ByVal value As Button)
            ppBackBtn = value
        End Set
    End Property



    ''' <summary>
    ''' プロパティセット【ppNewEntryBtn】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppNewEntryBtn</returns>
    ''' <remarks><para>作成情報：2015/08/11  yu.satoh
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropNewEntryBtn() As Button
        Get
            Return ppNewEntryBtn
        End Get
        Set(ByVal value As Button)
            ppNewEntryBtn = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【ppEntryBtn】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppEntryBtn</returns>
    ''' <remarks><para>作成情報：2015/08/11  yu.satoh
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropEntryBtn() As Button
        Get
            Return ppEntryBtn
        End Get
        Set(ByVal value As Button)
            ppEntryBtn = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【ppFromYearUp】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppFromYearUp</returns>
    ''' <remarks><para>作成情報：2015/08/11  yu.satoh
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropFromYearUp() As String
        Get
            Return ppFromYearUp
        End Get
        Set(ByVal value As String)
            ppFromYearUp = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【ppFromMonthUp】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppFromMonthUp</returns>
    ''' <remarks><para>作成情報：2015/08/11  yu.satoh
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropFromMonthUp() As String
        Get
            Return ppFromMonthUp
        End Get
        Set(ByVal value As String)
            ppFromMonthUp = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【ppToYearUp】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppToYearUp</returns>
    ''' <remarks><para>作成情報：2015/08/11  yu.satoh
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropToYearUp() As String
        Get
            Return ppToYearUp
        End Get
        Set(ByVal value As String)
            ppToYearUp = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【ppToMonthUp】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppToMonthUp</returns>
    ''' <remarks><para>作成情報：2015/08/11  yu.satoh
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropToMonthUp() As String
        Get
            Return ppToMonthUp
        End Get
        Set(ByVal value As String)
            ppToMonthUp = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【ppMaxBunruiCd】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppMaxBunruiCd</returns>
    ''' <remarks><para>作成情報：2015/08/11  yu.satoh
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropMaxBunruiCd() As String
        Get
            Return ppMaxBunruiCd
        End Get
        Set(ByVal value As String)
            ppMaxBunruiCd = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【ppMaxFutaiBunruiCd】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppMaxFutaiBunruiCd</returns>
    ''' <remarks><para>作成情報：2015/08/11  yu.satoh
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropMaxFutaiBunruiCd() As Integer
        Get
            Return ppMaxFutaiBunruiCd
        End Get
        Set(ByVal value As Integer)
            ppMaxFutaiBunruiCd = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【ppUpdBunruiCd】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppUpdBunruiCd</returns>
    ''' <remarks><para>作成情報：2015/08/11  yu.satoh
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropUpdBunruiCd() As Integer
        Get
            Return ppUpdBunruiCd
        End Get
        Set(ByVal value As Integer)
            ppUpdBunruiCd = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【EXTM付帯設備分類（表示用）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtFutaiMstDsp</returns>
    ''' <remarks><para>作成情報：2015.10.26 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtFutaiMstDsp() As DataTable
        Get
            Return ppDtFutaiMstDsp
        End Get
        Set(ByVal value As DataTable)
            ppDtFutaiMstDsp = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【EXTM付帯設備分類（退避用）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtFutaiMstEscp</returns>
    ''' <remarks><para>作成情報：2015.10.26 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtFutaiMstEscp() As DataTable
        Get
            Return ppDtFutaiMstEscp
        End Get
        Set(ByVal value As DataTable)
            ppDtFutaiMstEscp = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【ppBunruiCd（表示判定用）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppBunruiCd</returns>
    ''' <remarks><para>作成情報：2015.10.26 h.hagiawara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropBunruiCd() As String
        Get
            Return ppBunruiCd
        End Get
        Set(ByVal value As String)
            ppBunruiCd = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【ppBunruiCd（表示判定用）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppRow</returns>
    ''' <remarks><para>作成情報：2015.10.26 h.hagiawara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropRow() As Integer
        Get
            Return ppRow
        End Get
        Set(ByVal value As Integer)
            ppRow = value
        End Set
    End Property


    ''' <summary>
    ''' プロパティセット【期間ＦＲＯＭ（利用料用コピー元設定）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppstrKikan_From</returns>
    ''' <remarks><para>作成情報：2015.12.25 h.hagiawara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropstrKikan_From() As String
        Get
            Return ppstrKikan_From
        End Get
        Set(ByVal value As String)
            ppstrKikan_From = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【期間ＴＯ（利用料用コピー元設定）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppstrKikan_To</returns>
    ''' <remarks><para>作成情報：2015.12.25 h.hagiawara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropstrKikan_To() As String
        Get
            Return ppstrKikan_To
        End Get
        Set(ByVal value As String)
            ppstrKikan_To = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【施設区分（利用料用コピー元設定）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppstrShisetu_kbn</returns>
    ''' <remarks><para>作成情報：2015.12.25 h.hagiawara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropstrShisetu_kbn() As String
        Get
            Return ppstrShisetu_kbn
        End Get
        Set(ByVal value As String)
            ppstrShisetu_kbn = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【分類ＣＤ（利用料用コピー元設定）】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppstrBunrui_Cd</returns>
    ''' <remarks><para>作成情報：2015.12.25 h.hagiawara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropstrBunrui_Cd() As String
        Get
            Return ppstrBunrui_Cd
        End Get
        Set(ByVal value As String)
            ppstrBunrui_Cd = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【利用料用科目マスタ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtCopyKamokuMst</returns>
    ''' <remarks><para>作成情報：2015.12.25 h.hagiawara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtCopyKamokuMst() As DataTable
        Get
            Return ppDtCopyKamokuMst
        End Get
        Set(ByVal value As DataTable)
            ppDtCopyKamokuMst = value
        End Set
    End Property

    ' 2016.06.24 ADD START↓ h.hagiwara コンボリスト設定方法変更対応
    ''' <summary>
    ''' プロパティセット【EXTM科目マスタ(細目)】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtSaimoku</returns>
    ''' <remarks><para>作成情報：2016.06.24 h.hagiwara 
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtSaimoku() As DataTable
        Get
            Return ppDtSaimoku
        End Get
        Set(ByVal value As DataTable)
            ppDtSaimoku = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【EXTM科目マスタ(内訳)】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtUchiwake</returns>
    ''' <remarks><para>作成情報：2016.06.24 h.hagiwara 
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtUchiwake() As DataTable
        Get
            Return ppDtUchiwake
        End Get
        Set(ByVal value As DataTable)
            ppDtUchiwake = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【EXTM科目マスタ(詳細)】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtShosai</returns>
    ''' <remarks><para>作成情報：2016.06.24 h.hagiwara 
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtShosai() As DataTable
        Get
            Return ppDtShosai
        End Get
        Set(ByVal value As DataTable)
            ppDtShosai = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【検索用勘定科目】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppstrSelKAnjo</returns>
    ''' <remarks><para>作成情報：2016.06.24 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrSelKanjo() As String
        Get
            Return ppstrSelKAnjo
        End Get
        Set(ByVal value As String)
            ppstrSelKAnjo = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【検索用細目】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppFromMonthUp</returns>
    ''' <remarks><para>作成情報：2016.06.24 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrSelSaimoku() As String
        Get
            Return ppstrSelSaimoku
        End Get
        Set(ByVal value As String)
            ppstrSelSaimoku = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【検索用内訳】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppFromMonthUp</returns>
    ''' <remarks><para>作成情報：2016.06.24 h.hagiwara
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrSelUchiwake() As String
        Get
            Return ppstrSelUchiwake
        End Get
        Set(ByVal value As String)
            ppstrSelUchiwake = value
        End Set
    End Property

    ' 2016.06.24 ADD END↑ h.hagiwara コンボリスト設定方法変更対応

End Class
