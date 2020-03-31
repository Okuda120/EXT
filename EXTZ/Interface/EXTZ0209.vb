Imports System.Configuration
Imports Common
Imports CommonEXT
Imports EXTZ
Imports FarPoint.Win.Spread
Imports FarPoint.Win.Spread.Model

''' <summary>
''' EXTZ0209
''' </summary>
''' <remarks>付帯設備登録／詳細
''' <para>作成情報：2015/09/01 k.machida
''' <p>改訂情報:</p>
''' </para></remarks>
Public Class EXTZ0209
    ' --- 2019/06/11 軽減税率対応 Start E.Okuda@Compass ---
    ' 列番号定数化
    Private Const COL_FUTAI_FAUTAI_NM As Integer = 0
    Private Const COL_FUTAI_FAUTAI_CD As Integer = 1
    Private Const COL_FUTAI_TANKA As Integer = 2
    Private Const COL_FUTAI_TANI As Integer = 3
    Private Const COL_FUTAI_BUNRUI_NM As Integer = 4
    Private Const COL_FUTAI_BUNRUI_CD As Integer = 5
    Private Const COL_FUTAI_NOTAX_FLG As Integer = 6
    ' --- 2020/03/24 税区分追加対応 Start E.Okuda@Compass ---
    Private Const COL_FUTAI_TAX_KBN As Integer = 7
    Private Const COL_FUTAI_ZEIRITSU As Integer = 8
    Private Const COL_FUTAI_SHUKEI_GRP As Integer = 9
    Private Const COL_FUTAI_KAMOKU_CD As Integer = 10
    Private Const COL_FUTAI_SAIMOKU_CD As Integer = 11
    Private Const COL_FUTAI_UCHI_CD As Integer = 12
    Private Const COL_FUTAI_SHOSAI_CD As Integer = 13
    Private Const COL_FUTAI_KARIKAMOKU_CD As Integer = 14
    Private Const COL_FUTAI_KARI_SAIMOKU_CD As Integer = 15
    Private Const COL_FUTAI_KARI_UCHI_CD As Integer = 16
    Private Const COL_FUTAI_KARI_SHOSAI_CD As Integer = 17
    Private Const COL_FUTAI_KAMOKU_NM As Integer = 18
    Private Const COL_FUTAI_SAIMOKU_NM As Integer = 19
    Private Const COL_FUTAI_UCHI_NM As Integer = 20
    Private Const COL_FUTAI_SHOSAI_NM As Integer = 21

    'Private Const COL_FUTAI_ZEIRITSU As Integer = 7
    'Private Const COL_FUTAI_SHUKEI_GRP As Integer = 8
    'Private Const COL_FUTAI_KAMOKU_CD As Integer = 9
    'Private Const COL_FUTAI_SAIMOKU_CD As Integer = 10
    'Private Const COL_FUTAI_UCHI_CD As Integer = 11
    'Private Const COL_FUTAI_SHOSAI_CD As Integer = 12
    'Private Const COL_FUTAI_KARIKAMOKU_CD As Integer = 13
    'Private Const COL_FUTAI_KARI_SAIMOKU_CD As Integer = 14
    'Private Const COL_FUTAI_KARI_UCHI_CD As Integer = 15
    'Private Const COL_FUTAI_KARI_SHOSAI_CD As Integer = 16
    'Private Const COL_FUTAI_KAMOKU_NM As Integer = 17
    'Private Const COL_FUTAI_SAIMOKU_NM As Integer = 18
    'Private Const COL_FUTAI_UCHI_NM As Integer = 19
    'Private Const COL_FUTAI_SHOSAI_NM As Integer = 20
    ' --- 2020/03/24 税区分追加対応 End E.Okuda@Compass ---

    Private Const COL_FUTAI_SELECTED_FUTAI_BUNRUI_NM As Integer = 0
    Private Const COL_FUTAI_SELECTED_FUTAI_NM As Integer = 1
    Private Const COL_FUTAI_SELECTED_FUTAI_TANKA As Integer = 2
    Private Const COL_FUTAI_SELECTED_FUTAI_SU As Integer = 3
    Private Const COL_FUTAI_SELECTED_FUTAI_TANI As Integer = 4
    Private Const COL_FUTAI_SELECTED_FUTAI_SHOKEI As Integer = 5
    Private Const COL_FUTAI_SELECTED_FUTAI_CHOSEI As Integer = 6
    Private Const COL_FUTAI_SELECTED_FUTAI_KIN As Integer = 7
    Private Const COL_FUTAI_SELECTED_NOTAX_FLG As Integer = 8

    ' --- 2020/03/24 税区分追加対応 Start E.Okuda@Compass ---
    Private Const COL_FUTAI_SELECTED_TAX_KBN As Integer = 9
    Private Const COL_FUTAI_SELECTED_ZEIRITSU As Integer = 10
    Private Const COL_FUTAI_SELECTED_FUTAI_KIN_TAX As Integer = 11
    Private Const COL_FUTAI_SELECTED_FUTAI_BIKO As Integer = 12
    Private Const COL_FUTAI_SELECTED_FUTAI_BUNRUI_CD As Integer = 13
    Private Const COL_FUTAI_SELECTED_FUTAI_CD As Integer = 14
    Private Const COL_FUTAI_SELECTED_NOTAX_FLG_HIDDEN As Integer = 15
    Private Const COL_FUTAI_SELECTED_SHUKEI_GRP As Integer = 16
    Private Const COL_FUTAI_SELECTED_KAMOKU_CD As Integer = 17
    Private Const COL_FUTAI_SELECTED_SAIMOKU_CD As Integer = 18
    Private Const COL_FUTAI_SELECTED_UCHI_CD As Integer = 19
    Private Const COL_FUTAI_SELECTED_SHOSAI_CD As Integer = 20
    Private Const COL_FUTAI_SELECTED_KARIKAMOKU As Integer = 21
    Private Const COL_FUTAI_SELECTED_KARI_SAIMOKU_CD As Integer = 22
    Private Const COL_FUTAI_SELECTED_KARI_UCHI_CD As Integer = 23
    Private Const COL_FUTAI_SELECTED_KARI_SHOSAI_CD As Integer = 24
    Private Const COL_FUTAI_SELECTED_KAMOKU_NM As Integer = 25
    Private Const COL_FUTAI_SELECTED_SAIMOKU_NM As Integer = 26
    Private Const COL_FUTAI_SELECTED_UCHI_NM As Integer = 27
    Private Const COL_FUTAI_SELECTED_SHOSAI_NM As Integer = 28
    Private Const COL_FUTAI_SELECTED_DUMMY As Integer = 29

    'Private Const COL_FUTAI_SELECTED_ZEIRITSU As Integer = 9
    'Private Const COL_FUTAI_SELECTED_FUTAI_KIN_TAX As Integer = 10
    'Private Const COL_FUTAI_SELECTED_FUTAI_BIKO As Integer = 11
    'Private Const COL_FUTAI_SELECTED_FUTAI_BUNRUI_CD As Integer = 12
    'Private Const COL_FUTAI_SELECTED_FUTAI_CD As Integer = 13
    'Private Const COL_FUTAI_SELECTED_NOTAX_FLG_HIDDEN As Integer = 14
    'Private Const COL_FUTAI_SELECTED_SHUKEI_GRP As Integer = 15
    'Private Const COL_FUTAI_SELECTED_KAMOKU_CD As Integer = 16
    'Private Const COL_FUTAI_SELECTED_SAIMOKU_CD As Integer = 17
    'Private Const COL_FUTAI_SELECTED_UCHI_CD As Integer = 18
    'Private Const COL_FUTAI_SELECTED_SHOSAI_CD As Integer = 19
    'Private Const COL_FUTAI_SELECTED_KARIKAMOKU As Integer = 20
    'Private Const COL_FUTAI_SELECTED_KARI_SAIMOKU_CD As Integer = 21
    'Private Const COL_FUTAI_SELECTED_KARI_UCHI_CD As Integer = 22
    'Private Const COL_FUTAI_SELECTED_KARI_SHOSAI_CD As Integer = 23
    'Private Const COL_FUTAI_SELECTED_KAMOKU_NM As Integer = 24
    'Private Const COL_FUTAI_SELECTED_SAIMOKU_NM As Integer = 25
    'Private Const COL_FUTAI_SELECTED_UCHI_NM As Integer = 26
    'Private Const COL_FUTAI_SELECTED_SHOSAI_NM As Integer = 27
    'Private Const COL_FUTAI_SELECTED_DUMMY As Integer = 28
    ' --- 2020/03/24 税区分追加対応 End E.Okuda@Compass ---

    ' 税区分
    Private Const DISP_UCHIZEI As String = "内税"
    Private Const VAL_UCHIZEI As String = "1"
    Private Const DISP_SOTOZEI As String = "外税"
    Private Const VAL_SOTOZEI As String = "0"

    ' --- 2019/06/11 軽減税率対応 End E.Okuda@Compass ---


    Private commonValidate As New CommonValidation      '共通ロジッククラス
    Private commonLogicEXT As New CommonLogicEXT      '共通ロジッククラス
    Public dataEXTZ0209 As New DataEXTZ0209     'データクラス
    Public logicEXTZ0209 As New LogicEXTZ0209     'データクラス
    Private ppBlnChangeFlg As Boolean

    ''' <summary>
    ''' 初期処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub EXTZ0209_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.lblYoyakuNo.Text = dataEXTZ0209.PropStrYoyakuNo
        Me.lblSaiji.Text = dataEXTZ0209.PropStrSaijiNm
        Me.lblRiyobi.Text = dataEXTZ0209.PropStrRiyobiDisp

        Dim index As Integer = 0
        logicEXTZ0209.GetBunrui(dataEXTZ0209)
        Me.cmbBunrui.DataSource = dataEXTZ0209.PropFutaiBunruiTable
        Me.cmbBunrui.DisplayMember = "bunrui_nm"
        Me.cmbBunrui.ValueMember = "bunrui_cd"

        dataEXTZ0209.PropVwFutaiSelecSheet = Me.fbFutaiSelected                     ' 2015.12.01 ADD h.hagiwara

        logicEXTZ0209.GetFutai(dataEXTZ0209)
        selectBunrui(dataEXTZ0209.PropFutaiBunruiTable.Rows(0)("bunrui_cd"))
        setDetailToSpread()
        updateTotal()

        ' 背景色設定
        Me.BackColor = commonLogicEXT.SetFormBackColor(CommonEXT.PropConfigrationFlg)
        If CommonEXT.PropConfigrationFlg = "1" Then
            ' 検証機の場合には背景イメージも表示しない
            Me.BackgroundImage = Nothing
        End If

    End Sub

    ''' <summary>
    ''' 明細情報をスプレッドに表示する
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub setDetailToSpread()
        If dataEXTZ0209.PropFutaiDetailTable Is Nothing Then
            Exit Sub
        End If
        Dim sheet As FarPoint.Win.Spread.SheetView = Me.fbFutaiSelected.ActiveSheet
        Dim table As DataTable = dataEXTZ0209.PropFutaiDetailTable
        ' 2015.12.21 ADD START↓ h.hagiwara
        Dim numCellKeijyo As New FarPoint.Win.Spread.CellType.NumberCellType()
        numCellKeijyo.MaximumValue = 99999999
        numCellKeijyo.ShowSeparator = True
        numCellKeijyo.DecimalPlaces = 0
        ' 2015.12.21 ADD END↑ h.hagiwara
        sheet.RowCount = table.Rows.Count
        ' --- 2019/06/03 軽減税率対応 Start E.Okuda@Compass ---
        ' 【税率】、【内税外税】列追加

        'sheet.ColumnCount = 25        ' 2015.11.13 ADD h.hagiwara
        'sheet.ColumnCount = 26        ' 2015.11.13 ADD h.hagiwara
        ' --- 2020/03/24 税区分追加対応 Start E.Okuda@Compass ---
        'sheet.ColumnCount = 29
        sheet.ColumnCount = 30

        Dim idxTaxKbn As Integer

        ' 設定ファイルから税区分名称を取得
        Dim arySpreadTitle As Array
        arySpreadTitle = Split(ConfigurationManager.AppSettings("TaxMstItemNm"), ",")

        ' --- 2020/03/24 税区分追加対応 End E.Okuda@Compass ---


        'カラム設定(コード列の非表示)

        Const COLUMN_HIDDEN_START_POS As Integer = COL_FUTAI_SELECTED_FUTAI_BUNRUI_CD

        Dim intLoopCnt As Integer
        For intLoopCnt = COLUMN_HIDDEN_START_POS To sheet.ColumnCount - 1
            sheet.Columns(intLoopCnt).Visible = False
        Next
        'sheet.Columns(9).Visible = False
        'sheet.Columns(10).Visible = False
        'sheet.Columns(11).Visible = False
        'sheet.Columns(12).Visible = False
        'sheet.Columns(13).Visible = False
        'sheet.Columns(14).Visible = False
        'sheet.Columns(15).Visible = False
        'sheet.Columns(16).Visible = False
        'sheet.Columns(17).Visible = False
        'sheet.Columns(18).Visible = False
        'sheet.Columns(19).Visible = False
        'sheet.Columns(20).Visible = False
        'sheet.Columns(21).Visible = False
        'sheet.Columns(22).Visible = False
        'sheet.Columns(23).Visible = False
        'sheet.Columns(24).Visible = False
        'sheet.Columns(25).Visible = False         ' 2015/11/13 ADD h.hagwiara

        ' --- 2019/06/03 軽減税率対応 End E.Okuda@Compass ---

        Dim index As New Integer
        Dim row As DataRow
        index = 0
        Do While index < table.Rows.Count
            row = table.Rows(index)
            ' --- 2019/06/03 軽減税率対応 Start E.Okuda@Compass ---
            sheet.Cells(index, COL_FUTAI_SELECTED_FUTAI_BUNRUI_NM).Locked = True
            sheet.Cells(index, COL_FUTAI_SELECTED_FUTAI_NM).Locked = True
            sheet.Cells(index, COL_FUTAI_SELECTED_FUTAI_TANKA).Locked = True
            sheet.Cells(index, COL_FUTAI_SELECTED_FUTAI_TANI).Locked = True
            sheet.Cells(index, COL_FUTAI_SELECTED_FUTAI_SHOKEI).Locked = True
            sheet.Cells(index, COL_FUTAI_SELECTED_FUTAI_KIN).Locked = True
            sheet.Cells(index, COL_FUTAI_SELECTED_NOTAX_FLG).Locked = True
            sheet.Cells(index, COL_FUTAI_SELECTED_ZEIRITSU).Locked = True
            sheet.Cells(index, COL_FUTAI_SELECTED_FUTAI_KIN_TAX).Locked = True

            sheet.Cells(index, COL_FUTAI_SELECTED_FUTAI_BUNRUI_NM).Value = row("futai_bunrui_nm")
            sheet.Cells(index, COL_FUTAI_SELECTED_FUTAI_NM).Value = row("futai_nm")
            sheet.Cells(index, COL_FUTAI_SELECTED_FUTAI_TANKA).Value = row("futai_tanka")
            sheet.Cells(index, COL_FUTAI_SELECTED_FUTAI_SU).Value = row("futai_su")
            sheet.Cells(index, COL_FUTAI_SELECTED_FUTAI_TANI).Value = row("futai_tani")
            sheet.Cells(index, COL_FUTAI_SELECTED_FUTAI_SHOKEI).Value = row("futai_shokei")
            sheet.Cells(index, COL_FUTAI_SELECTED_FUTAI_CHOSEI).CellType = numCellKeijyo                      ' 2015.12.21 ADD h.hagiwara
            sheet.Cells(index, COL_FUTAI_SELECTED_FUTAI_CHOSEI).Value = row("futai_chosei")
            sheet.Cells(index, COL_FUTAI_SELECTED_FUTAI_KIN).Value = row("futai_kin")
            ' 税区分表記変更
            If row("notax_flg") = VAL_SOTOZEI Then
                sheet.Cells(index, COL_FUTAI_SELECTED_NOTAX_FLG).Value = DISP_SOTOZEI
            Else
                sheet.Cells(index, COL_FUTAI_SELECTED_NOTAX_FLG).Value = DISP_UCHIZEI
            End If

            ' sheet.Cells(index, COL_FUTAI_SELECTED_NOTAX_FLG).Value = row("notax_flg")

            ' --- 2020/03/24 税区分追加対応 Start E.Okuda@Compass ---
            sheet.Cells(index, COL_FUTAI_SELECTED_TAX_KBN).Locked = True
            If Integer.TryParse(row("tax_kbn"), idxTaxKbn) Then
                sheet.Cells(index, COL_FUTAI_SELECTED_TAX_KBN).Value = arySpreadTitle(idxTaxKbn - 1)
            End If

            sheet.Cells(index, COL_FUTAI_SELECTED_ZEIRITSU).Value = row("tax_rate")
            'sheet.Cells(index, COL_FUTAI_SELECTED_ZEIRITSU).Value = row("zeiritsu")
            ' --- 2020/03/24 税区分追加対応 End E.Okuda@Compass ---

            ' 合計額(税込)
            If row("notax_flg") = VAL_SOTOZEI Then
                sheet.Cells(index, COL_FUTAI_SELECTED_FUTAI_KIN_TAX).Value =
                    Long.Parse(row("futai_kin")) + Long.Parse(row("futai_kin")) * Integer.Parse(sheet.Cells(index, COL_FUTAI_SELECTED_ZEIRITSU).Value) / 100
            Else
                sheet.Cells(index, COL_FUTAI_SELECTED_FUTAI_KIN_TAX).Value = row("futai_kin")
            End If

            sheet.Cells(index, COL_FUTAI_SELECTED_FUTAI_BIKO).Value = row("futai_biko")
            sheet.Cells(index, COL_FUTAI_SELECTED_FUTAI_BUNRUI_CD).Value = row("futai_bunrui_cd")
            sheet.Cells(index, COL_FUTAI_SELECTED_FUTAI_CD).Value = row("futai_cd")
            sheet.Cells(index, COL_FUTAI_SELECTED_NOTAX_FLG_HIDDEN).Value = row("notax_flg")
            sheet.Cells(index, COL_FUTAI_SELECTED_SHUKEI_GRP).Value = row("shukei_grp")
            sheet.Cells(index, COL_FUTAI_SELECTED_KAMOKU_CD).Value = row("kamoku_cd")
            sheet.Cells(index, COL_FUTAI_SELECTED_SAIMOKU_CD).Value = row("saimoku_cd")
            sheet.Cells(index, COL_FUTAI_SELECTED_UCHI_CD).Value = row("uchi_cd")
            sheet.Cells(index, COL_FUTAI_SELECTED_SHOSAI_CD).Value = row("shosai_cd")
            sheet.Cells(index, COL_FUTAI_SELECTED_KARIKAMOKU).Value = row("karikamoku_cd")
            sheet.Cells(index, COL_FUTAI_SELECTED_KARI_SAIMOKU_CD).Value = row("kari_saimoku_cd")
            sheet.Cells(index, COL_FUTAI_SELECTED_KARI_UCHI_CD).Value = row("kari_uchi_cd")
            sheet.Cells(index, COL_FUTAI_SELECTED_KARI_SHOSAI_CD).Value = row("kari_shosai_cd")
            sheet.Cells(index, COL_FUTAI_SELECTED_KAMOKU_NM).Value = row("kamoku_nm")
            sheet.Cells(index, COL_FUTAI_SELECTED_SAIMOKU_NM).Value = row("saimoku_nm")
            sheet.Cells(index, COL_FUTAI_SELECTED_UCHI_NM).Value = row("uchi_nm")
            sheet.Cells(index, COL_FUTAI_SELECTED_SHOSAI_NM).Value = row("shosai_nm")
            sheet.Cells(index, COL_FUTAI_SELECTED_DUMMY).Value = 0                        ' 2015/11/13 ADD h.hagwiara

            'sheet.Cells(index, 7).Value = row("futai_kin")
            'sheet.Cells(index, 8).Value = row("futai_biko")
            'sheet.Cells(index, 9).Value = row("futai_bunrui_cd")
            'sheet.Cells(index, 10).Value = row("futai_cd")
            'sheet.Cells(index, 11).Value = row("notax_flg")
            'sheet.Cells(index, 12).Value = row("shukei_grp")
            'sheet.Cells(index, 13).Value = row("kamoku_cd")
            'sheet.Cells(index, 14).Value = row("saimoku_cd")
            'sheet.Cells(index, 15).Value = row("uchi_cd")
            'sheet.Cells(index, 16).Value = row("shosai_cd")
            'sheet.Cells(index, 17).Value = row("karikamoku_cd")
            'sheet.Cells(index, 18).Value = row("kari_saimoku_cd")
            'sheet.Cells(index, 19).Value = row("kari_uchi_cd")
            'sheet.Cells(index, 20).Value = row("kari_shosai_cd")
            'sheet.Cells(index, 21).Value = row("kamoku_nm")
            'sheet.Cells(index, 22).Value = row("saimoku_nm")
            'sheet.Cells(index, 23).Value = row("uchi_nm")
            'sheet.Cells(index, 24).Value = row("shosai_nm")
            'sheet.Cells(index, 25).Value = 0                        ' 2015/11/13 ADD h.hagwiara
            ' --- 2019/06/03 軽減税率対応 End E.Okuda@Compass ---
            index = index + 1
        Loop
        'スクロールバーを必要な場合のみ表示させます
        'Me.fbFutaiSelected.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Never                        ' 2015.12.21 UPD h.hagiwara
        Me.fbFutaiSelected.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.AsNeeded                      ' 2015.12.21 UPD h.hagiwara
        Me.fbFutaiSelected.VerticalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.AsNeeded
    End Sub

    ''' <summary>
    ''' 分類マスタのデータセットに対してセレクトを行う
    ''' </summary>
    ''' <param name="strBunruiCd"></param>
    ''' <remarks></remarks>
    Private Sub selectBunrui(ByVal strBunruiCd As String)
        If dataEXTZ0209.PropFutaiMstTable Is Nothing Then
            Exit Sub
        End If
        Dim drFutai() As DataRow
        Dim strBunruiKey As String = "bunrui_cd = '" + strBunruiCd + "'"
        drFutai = dataEXTZ0209.PropFutaiMstTable.Select(strBunruiKey, "sort")
        Dim drBunrui() As DataRow
        drBunrui = dataEXTZ0209.PropFutaiBunruiTable.Select(strBunruiKey, "sort")

        Dim sheet As FarPoint.Win.Spread.SheetView = Me.fbFutai.ActiveSheet
        sheet.RowCount = drFutai.Count

        ' --- 2020/03/24 税区分追加対応 Start E.Okuda@Compass ---
        sheet.ColumnCount = 22
        ' --- 2019/06/11 軽減税率対応 Start E.Okuda@Compass ---
        'sheet.ColumnCount = 21
        'sheet.ColumnCount = 20
        ' --- 2020/03/24 税区分追加対応 End E.Okuda@Compass ---
        'カラム設定(コード列の非表示)
        ' 2019/06/11 Loopに変更　E.Okuda@Compass
        Dim intLoopCnt As Integer
        For intLoopCnt = COL_FUTAI_FAUTAI_CD To sheet.ColumnCount - 1
            sheet.Columns(intLoopCnt).Visible = False
        Next

        'sheet.Columns(1).Visible = False
        'sheet.Columns(2).Visible = False
        'sheet.Columns(3).Visible = False
        'sheet.Columns(4).Visible = False
        'sheet.Columns(5).Visible = False
        'sheet.Columns(6).Visible = False
        'sheet.Columns(7).Visible = False
        'sheet.Columns(8).Visible = False
        'sheet.Columns(9).Visible = False
        'sheet.Columns(10).Visible = False
        'sheet.Columns(11).Visible = False
        'sheet.Columns(12).Visible = False
        'sheet.Columns(13).Visible = False
        'sheet.Columns(14).Visible = False
        'sheet.Columns(15).Visible = False
        'sheet.Columns(16).Visible = False
        'sheet.Columns(17).Visible = False
        'sheet.Columns(18).Visible = False
        'sheet.Columns(19).Visible = False
        Dim index As Integer = 0
        Dim bunruiRow As DataRow = drBunrui(0)
        For Each row As DataRow In drFutai
            sheet.Cells(index, COL_FUTAI_FAUTAI_NM).Value = row("futai_nm")
            sheet.Cells(index, COL_FUTAI_FAUTAI_NM).Locked = True
            sheet.Cells(index, COL_FUTAI_FAUTAI_CD).Value = row("futai_cd")
            sheet.Cells(index, COL_FUTAI_TANKA).Value = row("tanka")
            sheet.Cells(index, COL_FUTAI_TANI).Value = row("tani")
            sheet.Cells(index, COL_FUTAI_BUNRUI_NM).Value = row("bunrui_nm")
            sheet.Cells(index, COL_FUTAI_BUNRUI_CD).Value = row("bunrui_cd")
            sheet.Cells(index, COL_FUTAI_NOTAX_FLG).Value = row("notax_flg")
            ' --- 2020/03/24 税区分追加対応 Start E.Okuda@Compass ---
            sheet.Cells(index, COL_FUTAI_TAX_KBN).Value = row("tax_kbn")
            ' --- 2020/03/24 税区分追加対応 End E.Okuda@Compass ---
            sheet.Cells(index, COL_FUTAI_ZEIRITSU).Value = row("tax_rate")
            sheet.Cells(index, COL_FUTAI_SHUKEI_GRP).Value = row("shukei_grp")
            sheet.Cells(index, COL_FUTAI_KAMOKU_CD).Value = row("kamoku_cd")
            sheet.Cells(index, COL_FUTAI_SAIMOKU_CD).Value = row("saimoku_cd")
            sheet.Cells(index, COL_FUTAI_UCHI_CD).Value = row("uchi_cd")
            sheet.Cells(index, COL_FUTAI_SHOSAI_CD).Value = row("shosai_cd")
            sheet.Cells(index, COL_FUTAI_KARIKAMOKU_CD).Value = row("karikamoku_cd")
            sheet.Cells(index, COL_FUTAI_KARI_SAIMOKU_CD).Value = row("kari_saimoku_cd")
            sheet.Cells(index, COL_FUTAI_KARI_UCHI_CD).Value = row("kari_uchi_cd")
            sheet.Cells(index, COL_FUTAI_KARI_SHOSAI_CD).Value = row("kari_shosai_cd")
            sheet.Cells(index, COL_FUTAI_KAMOKU_NM).Value = row("kamoku_nm")
            sheet.Cells(index, COL_FUTAI_SAIMOKU_NM).Value = row("saimoku_nm")
            sheet.Cells(index, COL_FUTAI_UCHI_NM).Value = row("uchi_nm")
            sheet.Cells(index, COL_FUTAI_SHOSAI_NM).Value = row("shosai_nm")

            'sheet.Cells(index, 0).Value = row("futai_nm")
            'sheet.Cells(index, 0).Locked = True
            'sheet.Cells(index, 1).Value = row("futai_cd")
            'sheet.Cells(index, 2).Value = row("tanka")
            'sheet.Cells(index, 3).Value = row("tani")
            'sheet.Cells(index, 4).Value = row("bunrui_nm")
            'sheet.Cells(index, 5).Value = row("bunrui_cd")
            'sheet.Cells(index, 6).Value = row("notax_flg")
            'sheet.Cells(index, 7).Value = row("shukei_grp")
            'sheet.Cells(index, 8).Value = row("kamoku_cd")
            'sheet.Cells(index, 9).Value = row("saimoku_cd")
            'sheet.Cells(index, 10).Value = row("uchi_cd")
            'sheet.Cells(index, 11).Value = row("shosai_cd")
            'sheet.Cells(index, 12).Value = row("karikamoku_cd")
            'sheet.Cells(index, 13).Value = row("kari_saimoku_cd")
            'sheet.Cells(index, 14).Value = row("kari_uchi_cd")
            'sheet.Cells(index, 15).Value = row("kari_shosai_cd")
            'sheet.Cells(index, 16).Value = row("kamoku_nm")
            'sheet.Cells(index, 17).Value = row("saimoku_nm")
            'sheet.Cells(index, 18).Value = row("uchi_nm")
            'sheet.Cells(index, 19).Value = row("shosai_nm")
            index = index + 1
        Next
        ' --- 2019/06/11 軽減税率対応 End E.Okuda@Compass ---

        'スクロールバーを必要な場合のみ表示させます
        Me.fbFutai.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Never
        Me.fbFutai.VerticalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.AsNeeded

    End Sub

    ''' <summary>
    ''' 分類選択時の処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub cmbBunrui_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbBunrui.SelectedIndexChanged
        selectBunrui(cmbBunrui.SelectedValue.ToString())
    End Sub

    ''' <summary>
    ''' 行コピー
    ''' </summary>
    ''' <param name="activeRow"></param>
    ''' <remarks></remarks>
    Private Sub copyrow(activeRow As Integer)
        Dim lastRow As Integer = fbFutaiSelected.ActiveSheet.RowCount
        ' 2015.12.21 ADD START↓ h.hagiwara
        Dim numCellKeijyo As New FarPoint.Win.Spread.CellType.NumberCellType()

        ' --- 2020/03/24 税区分追加対応 Start E.Okuda@Compass ---
        Dim idxTaxKbn As Integer

        ' 設定ファイルから税区分名称を取得
        Dim arySpreadTitle As Array
        arySpreadTitle = Split(ConfigurationManager.AppSettings("TaxMstItemNm"), ",")

        ' --- 2020/03/24 税区分追加対応 End E.Okuda@Compass ---

        numCellKeijyo.MaximumValue = 99999999
        numCellKeijyo.ShowSeparator = True
        numCellKeijyo.DecimalPlaces = 0
        ' 2015.12.21 ADD END↑ h.hagiwara

        '最終行に1行追加
        fbFutaiSelected.Sheets(0).AddUnboundRows(lastRow, 1)

        '設定
        ' --- 2019/06/06 軽減税率対応 Start E.Okuda@Compass ---
        ' 列番号定数化＆「税区分」、「税率」追加

        fbFutaiSelected.ActiveSheet.Cells(lastRow, COL_FUTAI_SELECTED_FUTAI_BUNRUI_NM).Locked = True
        fbFutaiSelected.ActiveSheet.Cells(lastRow, COL_FUTAI_SELECTED_FUTAI_NM).Locked = True
        fbFutaiSelected.ActiveSheet.Cells(lastRow, COL_FUTAI_SELECTED_FUTAI_TANKA).Locked = True
        fbFutaiSelected.ActiveSheet.Cells(lastRow, COL_FUTAI_SELECTED_FUTAI_TANI).Locked = True
        fbFutaiSelected.ActiveSheet.Cells(lastRow, COL_FUTAI_SELECTED_FUTAI_SHOKEI).Locked = True
        fbFutaiSelected.ActiveSheet.Cells(lastRow, COL_FUTAI_SELECTED_FUTAI_KIN).Locked = True
        fbFutaiSelected.ActiveSheet.Cells(lastRow, COL_FUTAI_SELECTED_NOTAX_FLG).Locked = True
        fbFutaiSelected.ActiveSheet.Cells(lastRow, COL_FUTAI_SELECTED_FUTAI_KIN_TAX).Locked = True

        '選択された設備情報をコピー
        fbFutaiSelected.ActiveSheet.Cells(lastRow, COL_FUTAI_SELECTED_FUTAI_BUNRUI_NM).Value = fbFutai.ActiveSheet.Cells(activeRow, COL_FUTAI_BUNRUI_NM).Value 'ジャンル
        fbFutaiSelected.ActiveSheet.Cells(lastRow, COL_FUTAI_SELECTED_FUTAI_NM).Value = fbFutai.ActiveSheet.Cells(activeRow, COL_FUTAI_FAUTAI_NM).Value '設備名
        fbFutaiSelected.ActiveSheet.Cells(lastRow, COL_FUTAI_SELECTED_FUTAI_TANKA).Value = fbFutai.ActiveSheet.Cells(activeRow, COL_FUTAI_TANKA).Value '単価
        fbFutaiSelected.ActiveSheet.Cells(lastRow, COL_FUTAI_SELECTED_FUTAI_SU).Value = 1                                             '数量
        fbFutaiSelected.ActiveSheet.Cells(lastRow, COL_FUTAI_SELECTED_FUTAI_TANI).Value = fbFutai.ActiveSheet.Cells(activeRow, COL_FUTAI_TANI).Value '単位
        fbFutaiSelected.ActiveSheet.Cells(lastRow, COL_FUTAI_SELECTED_FUTAI_SHOKEI).Value = fbFutai.ActiveSheet.Cells(activeRow, COL_FUTAI_TANKA).Value '小計→ tani * su
        fbFutaiSelected.ActiveSheet.Cells(lastRow, COL_FUTAI_SELECTED_FUTAI_CHOSEI).CellType = numCellKeijyo                              ' 2015.12.21 ADD h.hagiwara
        fbFutaiSelected.ActiveSheet.Cells(lastRow, COL_FUTAI_SELECTED_FUTAI_CHOSEI).Value = 0                                             '調整額
        fbFutaiSelected.ActiveSheet.Cells(lastRow, COL_FUTAI_SELECTED_FUTAI_KIN).Value = fbFutai.ActiveSheet.Cells(activeRow, COL_FUTAI_TANKA).Value '合計→ tani * su - chosei
        ' 税区分表記変更
        If fbFutai.ActiveSheet.Cells(activeRow, COL_FUTAI_NOTAX_FLG).Value = VAL_SOTOZEI Then
            fbFutaiSelected.ActiveSheet.Cells(lastRow, COL_FUTAI_SELECTED_NOTAX_FLG).Value = DISP_SOTOZEI
        Else
            fbFutaiSelected.ActiveSheet.Cells(lastRow, COL_FUTAI_SELECTED_NOTAX_FLG).Value = DISP_UCHIZEI
        End If

        ' --- 2020/03/24 税区分追加対応 Start E.Okuda@Compass ---
        If Integer.TryParse(fbFutai.ActiveSheet.Cells(activeRow, COL_FUTAI_TAX_KBN).Value, idxTaxKbn) Then
            fbFutaiSelected.ActiveSheet.Cells(lastRow, COL_FUTAI_SELECTED_TAX_KBN).Value = arySpreadTitle(idxTaxKbn - 1)
        End If
        ' --- 2020/03/24 税区分追加対応 End E.Okuda@Compass ---

        fbFutaiSelected.ActiveSheet.Cells(lastRow, COL_FUTAI_SELECTED_ZEIRITSU).Value = fbFutai.ActiveSheet.Cells(activeRow, COL_FUTAI_ZEIRITSU).Value
        If fbFutai.ActiveSheet.Cells(activeRow, COL_FUTAI_NOTAX_FLG).Value = VAL_SOTOZEI Then
            fbFutaiSelected.ActiveSheet.Cells(lastRow, COL_FUTAI_SELECTED_FUTAI_KIN_TAX).Value = Long.Parse(fbFutai.ActiveSheet.Cells(activeRow, COL_FUTAI_TANKA).Value) +
                Long.Parse(fbFutai.ActiveSheet.Cells(activeRow, COL_FUTAI_TANKA).Value) * Long.Parse(fbFutai.ActiveSheet.Cells(activeRow, COL_FUTAI_ZEIRITSU).Value) / 100
        Else
            fbFutaiSelected.ActiveSheet.Cells(lastRow, COL_FUTAI_SELECTED_FUTAI_KIN_TAX).Value = fbFutai.ActiveSheet.Cells(activeRow, COL_FUTAI_TANKA).Value
        End If
        fbFutaiSelected.ActiveSheet.Cells(lastRow, COL_FUTAI_SELECTED_FUTAI_BIKO).Value = Nothing
        fbFutaiSelected.ActiveSheet.Cells(lastRow, COL_FUTAI_SELECTED_FUTAI_BUNRUI_CD).Value = fbFutai.ActiveSheet.Cells(activeRow, COL_FUTAI_BUNRUI_CD).Value '分類CD
        fbFutaiSelected.ActiveSheet.Cells(lastRow, COL_FUTAI_SELECTED_FUTAI_CD).Value = fbFutai.ActiveSheet.Cells(activeRow, COL_FUTAI_FAUTAI_CD).Value '付帯CD
        fbFutaiSelected.ActiveSheet.Cells(lastRow, COL_FUTAI_SELECTED_NOTAX_FLG_HIDDEN).Value = fbFutai.ActiveSheet.Cells(activeRow, COL_FUTAI_NOTAX_FLG).Value
        fbFutaiSelected.ActiveSheet.Cells(lastRow, COL_FUTAI_SELECTED_SHUKEI_GRP).Value = fbFutai.ActiveSheet.Cells(activeRow, COL_FUTAI_SHUKEI_GRP).Value
        fbFutaiSelected.ActiveSheet.Cells(lastRow, COL_FUTAI_SELECTED_KAMOKU_CD).Value = fbFutai.ActiveSheet.Cells(activeRow, COL_FUTAI_KAMOKU_CD).Value
        fbFutaiSelected.ActiveSheet.Cells(lastRow, COL_FUTAI_SELECTED_SAIMOKU_CD).Value = fbFutai.ActiveSheet.Cells(activeRow, COL_FUTAI_SAIMOKU_CD).Value
        fbFutaiSelected.ActiveSheet.Cells(lastRow, COL_FUTAI_SELECTED_UCHI_CD).Value = fbFutai.ActiveSheet.Cells(activeRow, COL_FUTAI_UCHI_CD).Value
        fbFutaiSelected.ActiveSheet.Cells(lastRow, COL_FUTAI_SELECTED_SHOSAI_CD).Value = fbFutai.ActiveSheet.Cells(activeRow, COL_FUTAI_SHOSAI_CD).Value
        fbFutaiSelected.ActiveSheet.Cells(lastRow, COL_FUTAI_SELECTED_KARIKAMOKU).Value = fbFutai.ActiveSheet.Cells(activeRow, COL_FUTAI_KARIKAMOKU_CD).Value
        fbFutaiSelected.ActiveSheet.Cells(lastRow, COL_FUTAI_SELECTED_KARI_SAIMOKU_CD).Value = fbFutai.ActiveSheet.Cells(activeRow, COL_FUTAI_KARI_SAIMOKU_CD).Value
        fbFutaiSelected.ActiveSheet.Cells(lastRow, COL_FUTAI_SELECTED_KARI_UCHI_CD).Value = fbFutai.ActiveSheet.Cells(activeRow, COL_FUTAI_KARI_UCHI_CD).Value
        fbFutaiSelected.ActiveSheet.Cells(lastRow, COL_FUTAI_SELECTED_KARI_SHOSAI_CD).Value = fbFutai.ActiveSheet.Cells(activeRow, COL_FUTAI_KARI_SHOSAI_CD).Value
        fbFutaiSelected.ActiveSheet.Cells(lastRow, COL_FUTAI_SELECTED_KAMOKU_NM).Value = fbFutai.ActiveSheet.Cells(activeRow, COL_FUTAI_KAMOKU_NM).Value
        fbFutaiSelected.ActiveSheet.Cells(lastRow, COL_FUTAI_SELECTED_SAIMOKU_NM).Value = fbFutai.ActiveSheet.Cells(activeRow, COL_FUTAI_SAIMOKU_NM).Value
        fbFutaiSelected.ActiveSheet.Cells(lastRow, COL_FUTAI_SELECTED_UCHI_NM).Value = fbFutai.ActiveSheet.Cells(activeRow, COL_FUTAI_UCHI_NM).Value
        fbFutaiSelected.ActiveSheet.Cells(lastRow, COL_FUTAI_SELECTED_SHOSAI_NM).Value = fbFutai.ActiveSheet.Cells(activeRow, COL_FUTAI_SHOSAI_NM).Value
        fbFutaiSelected.ActiveSheet.Cells(lastRow, COL_FUTAI_SELECTED_DUMMY).Value = 0                                           ' 2015/11/13 ADD h.hagiwara

        'fbFutaiSelected.ActiveSheet.Cells(lastRow, 0).Locked = True
        'fbFutaiSelected.ActiveSheet.Cells(lastRow, 1).Locked = True
        'fbFutaiSelected.ActiveSheet.Cells(lastRow, 2).Locked = True
        'fbFutaiSelected.ActiveSheet.Cells(lastRow, 4).Locked = True
        'fbFutaiSelected.ActiveSheet.Cells(lastRow, 5).Locked = True
        'fbFutaiSelected.ActiveSheet.Cells(lastRow, 7).Locked = True
        ''選択された設備情報をコピー
        'fbFutaiSelected.ActiveSheet.Cells(lastRow, 0).Value = fbFutai.ActiveSheet.Cells(activeRow, 4).Value 'ジャンル
        'fbFutaiSelected.ActiveSheet.Cells(lastRow, 1).Value = fbFutai.ActiveSheet.Cells(activeRow, 0).Value '設備名
        'fbFutaiSelected.ActiveSheet.Cells(lastRow, 2).Value = fbFutai.ActiveSheet.Cells(activeRow, 2).Value '単価
        'fbFutaiSelected.ActiveSheet.Cells(lastRow, 3).Value = 1                                             '数量
        'fbFutaiSelected.ActiveSheet.Cells(lastRow, 4).Value = fbFutai.ActiveSheet.Cells(activeRow, 3).Value '単位
        'fbFutaiSelected.ActiveSheet.Cells(lastRow, 5).Value = fbFutai.ActiveSheet.Cells(activeRow, 2).Value '小計→ tani * su
        'fbFutaiSelected.ActiveSheet.Cells(lastRow, 6).CellType = numCellKeijyo                              ' 2015.12.21 ADD h.hagiwara
        'fbFutaiSelected.ActiveSheet.Cells(lastRow, 6).Value = 0                                             '調整額
        'fbFutaiSelected.ActiveSheet.Cells(lastRow, 7).Value = fbFutai.ActiveSheet.Cells(activeRow, 2).Value '合計→ tani * su - chosei
        'fbFutaiSelected.ActiveSheet.Cells(lastRow, 8).Value = Nothing
        'fbFutaiSelected.ActiveSheet.Cells(lastRow, 9).Value = fbFutai.ActiveSheet.Cells(activeRow, 5).Value '分類CD
        'fbFutaiSelected.ActiveSheet.Cells(lastRow, 10).Value = fbFutai.ActiveSheet.Cells(activeRow, 1).Value '付帯CD
        'fbFutaiSelected.ActiveSheet.Cells(lastRow, 11).Value = fbFutai.ActiveSheet.Cells(activeRow, 6).Value
        'fbFutaiSelected.ActiveSheet.Cells(lastRow, 12).Value = fbFutai.ActiveSheet.Cells(activeRow, 7).Value
        'fbFutaiSelected.ActiveSheet.Cells(lastRow, 13).Value = fbFutai.ActiveSheet.Cells(activeRow, 8).Value
        'fbFutaiSelected.ActiveSheet.Cells(lastRow, 14).Value = fbFutai.ActiveSheet.Cells(activeRow, 9).Value
        'fbFutaiSelected.ActiveSheet.Cells(lastRow, 15).Value = fbFutai.ActiveSheet.Cells(activeRow, 10).Value
        'fbFutaiSelected.ActiveSheet.Cells(lastRow, 16).Value = fbFutai.ActiveSheet.Cells(activeRow, 11).Value
        'fbFutaiSelected.ActiveSheet.Cells(lastRow, 17).Value = fbFutai.ActiveSheet.Cells(activeRow, 12).Value
        'fbFutaiSelected.ActiveSheet.Cells(lastRow, 18).Value = fbFutai.ActiveSheet.Cells(activeRow, 13).Value
        'fbFutaiSelected.ActiveSheet.Cells(lastRow, 19).Value = fbFutai.ActiveSheet.Cells(activeRow, 14).Value
        'fbFutaiSelected.ActiveSheet.Cells(lastRow, 20).Value = fbFutai.ActiveSheet.Cells(activeRow, 15).Value
        'fbFutaiSelected.ActiveSheet.Cells(lastRow, 21).Value = fbFutai.ActiveSheet.Cells(activeRow, 16).Value
        'fbFutaiSelected.ActiveSheet.Cells(lastRow, 22).Value = fbFutai.ActiveSheet.Cells(activeRow, 17).Value
        'fbFutaiSelected.ActiveSheet.Cells(lastRow, 23).Value = fbFutai.ActiveSheet.Cells(activeRow, 18).Value
        'fbFutaiSelected.ActiveSheet.Cells(lastRow, 24).Value = fbFutai.ActiveSheet.Cells(activeRow, 19).Value
        'fbFutaiSelected.ActiveSheet.Cells(lastRow, 25).Value = 0                                           ' 2015/11/13 ADD h.hagiwara
        '付帯設備合計を更新
        updateTotal()
    End Sub

    ''' <summary>
    ''' 業削除
    ''' </summary>
    ''' <param name="activeRow"></param>
    ''' <remarks></remarks>
    Private Sub removerow(activeRow As Integer)
        Dim row As Integer = fbFutaiSelected.Sheets(0).ActiveRowIndex
        If row <> -1 Then
            fbFutaiSelected.Sheets(0).RemoveRows(row, 1)
        End If
        '付帯設備合計を更新
        updateTotal()
    End Sub

    ''' <summary>
    ''' 料金自動計算
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Public Sub ChangeFutaiKin(ByVal sender As Object, ByVal e As ChangeEventArgs) Handles fbFutaiSelected.Change
        ' --- 2019/06/06 軽減税率対応 Start E.Okuda@Compass ---
        ' 列番号定数化
        If COL_FUTAI_SELECTED_FUTAI_SU <> e.Column And COL_FUTAI_SELECTED_FUTAI_CHOSEI <> e.Column Then
            Exit Sub
        End If
        'If 3 <> e.Column And 6 <> e.Column Then
        '    Exit Sub
        'End If
        ' --- 2019/06/06 軽減税率対応 End E.Okuda@Compass ---

        Dim sheet As FarPoint.Win.Spread.SheetView = Me.fbFutaiSelected.ActiveSheet
        'Dim tanka As Integer                      ' 2015.12.21 UPD  h.hagiwara
        Dim tanka As Long                          ' 2015.12.21 UPD  h.hagiwara
        Dim su As Integer
        ' 2015.12.21 UPD START↓ h.hagiwara
        'Dim shokei As Integer
        'Dim chosei As Integer
        'Dim total As Integer
        Dim shokei As Long
        Dim chosei As Integer
        Dim total As Long
        ' 2015.12.21 UPD END↑ h.hagiwara

        ' --- 2019/06/06 軽減税率対応 Start E.Okuda@Compass ---
        ' 列番号定数化
        Dim strSu As String = sheet.Cells(e.Row, COL_FUTAI_SELECTED_FUTAI_SU).Value
        Dim strChosei As String = sheet.Cells(e.Row, COL_FUTAI_SELECTED_FUTAI_CHOSEI).Value
        'Dim strSu As String = sheet.Cells(e.Row, 3).Value
        'Dim strChosei As String = sheet.Cells(e.Row, 6).Value
        ' --- 2019/06/06 軽減税率対応 End E.Okuda@Compass ---

        If commonValidate.IsHalfNmb(strSu) = False Then
            MsgBox(String.Format(CommonDeclareEXT.E0003, "数量"), MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
        If commonValidate.IsHalfMNmb(strChosei) = False Then
            MsgBox(String.Format(CommonDeclareEXT.E0004, "調整金額"), MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
        If String.IsNullOrEmpty(strSu) = True Then
            strSu = 1
            ' --- 2019/06/06 軽減税率対応 Start E.Okuda@Compass ---
            ' 列番号定数化
            sheet.Cells(e.Row, COL_FUTAI_SELECTED_FUTAI_SU).Value = 1
            'sheet.Cells(e.Row, 3).Value = 1
            ' --- 2019/06/06 軽減税率対応 End E.Okuda@Compass ---
        End If
        If String.IsNullOrEmpty(strChosei) = True Then
            ' --- 2019/06/06 軽減税率対応 Start E.Okuda@Compass ---
            ' 列番号定数化
            sheet.Cells(e.Row, COL_FUTAI_SELECTED_FUTAI_CHOSEI).Value = 0
            'sheet.Cells(e.Row, 6).Value = 0
            ' --- 2019/06/06 軽減税率対応 End E.Okuda@Compass ---
            strChosei = 0
        End If
        ' --- 2019/06/06 軽減税率対応 Start E.Okuda@Compass ---
        ' 列番号定数化
        tanka = sheet.Cells(e.Row, COL_FUTAI_SELECTED_FUTAI_TANKA).Value
        'tanka = sheet.Cells(e.Row, 2).Value
        ' --- 2019/06/06 軽減税率対応 End E.Okuda@Compass ---
        su = Integer.Parse(strSu)
        shokei = tanka * su
        chosei = Integer.Parse(strChosei)
        total = shokei + chosei

        ' 2015.12.21 UPD START↓ h.hagiwara
        If total.ToString.Length > 11 Then
            MsgBox(String.Format(CommonDeclareEXT.E2049, "付帯設備"), MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
        ' 2015.12.21 UPD END↑ h.hagiwara

        ' --- 2019/06/06 軽減税率対応 Start E.Okuda@Compass ---
        ' 列番号定数化
        sheet.Cells(e.Row, COL_FUTAI_SELECTED_FUTAI_SHOKEI).Value = shokei
        sheet.Cells(e.Row, COL_FUTAI_SELECTED_FUTAI_KIN).Value = total
        ' 合計(税込)表示
        If sheet.Cells(e.Row, COL_FUTAI_SELECTED_NOTAX_FLG).Value = DISP_SOTOZEI Then
            sheet.Cells(e.Row, COL_FUTAI_SELECTED_FUTAI_KIN_TAX).Value = total + total * Long.Parse(sheet.Cells(e.Row, COL_FUTAI_SELECTED_ZEIRITSU).Value) / 100
        Else
            sheet.Cells(e.Row, COL_FUTAI_SELECTED_FUTAI_KIN_TAX).Value = total
        End If


        'sheet.Cells(e.Row, 5).Value = shokei
        'sheet.Cells(e.Row, 7).Value = total


        ' --- 2019/06/06 軽減税率対応 End E.Okuda@Compass ---

        '合計再計算
        updateTotal()
    End Sub

    ''' <summary>
    ''' ＞ボタン押下処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        copyrow(fbFutai.ActiveSheet.ActiveRowIndex)
    End Sub

    ''' <summary>
    ''' ＞＞ボタン押下処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnAddAll_Click(sender As Object, e As EventArgs) Handles btnAddAll.Click
        Dim lastRow As Integer = fbFutai.ActiveSheet.GetLastNonEmptyRow(FarPoint.Win.Spread.NonEmptyItemFlag.Data)
        For i = 0 To lastRow
            copyrow(i)
        Next
    End Sub

    ''' <summary>
    ''' ＜ボタン押下処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnDel_Click(sender As Object, e As EventArgs) Handles btnDel.Click
        removerow(fbFutai.ActiveSheet.ActiveRowIndex)
    End Sub

    ''' <summary>
    ''' ＜＜ボタン押下処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnDelAll_Click(sender As Object, e As EventArgs) Handles btnDelAll.Click
        Dim lastRow As Integer = fbFutaiSelected.ActiveSheet.GetLastNonEmptyRow(FarPoint.Win.Spread.NonEmptyItemFlag.Data)
        For i = 0 To lastRow
            removerow(0)
        Next
        '付帯設備合計を更新
        updateTotal()
    End Sub

    ''' <summary>
    ''' 合計の再計算
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub updateTotal()
        ' --- 2019/06/06 軽減税率対応 Start E.Okuda@Compass ---
        ' 列番号定数化
        fbFutaiSelected.ActiveSheet.ColumnFooter.SetAggregationType(0, COL_FUTAI_SELECTED_FUTAI_KIN, FarPoint.Win.Spread.Model.AggregationType.Sum)
        lblTotal.Text = Format(fbFutaiSelected.Sheets(0).ColumnFooter.Cells(0, COL_FUTAI_SELECTED_FUTAI_KIN).Value, "#,#")
        'fbFutaiSelected.ActiveSheet.ColumnFooter.SetAggregationType(0, 7, FarPoint.Win.Spread.Model.AggregationType.Sum)
        'lblTotal.Text = Format(fbFutaiSelected.Sheets(0).ColumnFooter.Cells(0, 7).Value, "#,#")
        fbFutaiSelected.ActiveSheet.ColumnFooter.SetAggregationType(0, COL_FUTAI_SELECTED_FUTAI_KIN_TAX, FarPoint.Win.Spread.Model.AggregationType.Sum)
        lblTotalTax.Text = Format(fbFutaiSelected.Sheets(0).ColumnFooter.Cells(0, COL_FUTAI_SELECTED_FUTAI_KIN_TAX).Value, "#,#")
        ' --- 2019/06/06 軽減税率対応 End E.Okuda@Compass ---
    End Sub

    ''' <summary>
    ''' 入力完了処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnComplate_Click(sender As Object, e As EventArgs) Handles btnComplate.Click
        'input check
        Dim sheet As FarPoint.Win.Spread.SheetView = Me.fbFutaiSelected.ActiveSheet
        Dim index As New Integer
        Dim newRow As DataRow
        Dim table As DataTable = dataEXTZ0209.PropFutaiDetailTable.clone
        Dim futaiTotalNm As String = ""
        Dim isFirst As Boolean = True

        ' --- 2020/03/25 税区分追加対応 Start E.Okuda@Compass ---
        Dim i As Integer
        ' 設定ファイルから税区分名称を取得
        Dim arySpreadTitle As Array
        arySpreadTitle = Split(ConfigurationManager.AppSettings("TaxMstItemNm"), ",")
        ' --- 2020/03/25 税区分追加対応 End E.Okuda@Compass ---

        ' 2015.12.01 ADD START↓ h.hagiwara チェック追加(重複チェックも含む)
        If logicEXTZ0209.InputCheck(dataEXTZ0209) = False Then
            'メッセージを出力 
            MsgBox(puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
        ' 2015.12.01 ADD END↑ h.hagiwara チェック追加(重複チェックも含む)

        ' DBレプリケーション
        If commonLogicEXT.CheckDBCondition() = False Then
            'メッセージを出力 
            MsgBox(CommonDeclare.puErrMsg, MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If

        ' 消費税率取得  2015/11/13 ADD h.hagiwara
        'Dim intTaxkin As Integer                                 ' 2015.12.21 UPD h.hagiwara
        'Dim intTotalTaxkin As Integer                            ' 2015.12.21 UPD h.hagiwara
        Dim intTaxkin As Long                                     ' 2015.12.21 UPD h.hagiwara
        Dim intTotalTaxkin As Long                                ' 2015.12.21 UPD h.hagiwara

        ' --- 2020/03/24 税区分追加対応 Start E.Okuda@Compass ---
        ' 不要の為、コメントアウト
        'Dim dblTaxritu As Double
        'dblTaxritu = logicEXTZ0209.GetTax(dataEXTZ0209.PropStrRiyobi)
        ' --- 2020/03/24 税区分追加対応 End E.Okuda@Compass ---

        'クリア
        table.Clear()
        index = 0
        Do While index < sheet.Rows.Count
            newRow = table.NewRow()

            ' --- 2019/06/11 軽減税率対応 Start E.Okuda@Compass ---
            ' 列番号定数化
            newRow("yoyaku_dt") = dataEXTZ0209.PropStrRiyobi
            newRow("futai_bunrui_cd") = sheet.Cells(index, COL_FUTAI_SELECTED_FUTAI_BUNRUI_CD).Value
            newRow("futai_cd") = sheet.Cells(index, COL_FUTAI_SELECTED_FUTAI_CD).Value
            newRow("futai_tanka") = sheet.Cells(index, COL_FUTAI_SELECTED_FUTAI_TANKA).Value
            newRow("futai_su") = sheet.Cells(index, COL_FUTAI_SELECTED_FUTAI_SU).Value
            newRow("futai_shokei") = sheet.Cells(index, COL_FUTAI_SELECTED_FUTAI_SHOKEI).Value
            newRow("futai_chosei") = sheet.Cells(index, COL_FUTAI_SELECTED_FUTAI_CHOSEI).Value
            newRow("futai_kin") = sheet.Cells(index, COL_FUTAI_SELECTED_FUTAI_KIN).Value
            newRow("futai_biko") = commonLogicEXT.DbNothingToNull(sheet.Cells(index, COL_FUTAI_SELECTED_FUTAI_BIKO).Value)
            newRow("futai_nm") = sheet.Cells(index, COL_FUTAI_SELECTED_FUTAI_NM).Value
            newRow("futai_tani") = sheet.Cells(index, COL_FUTAI_SELECTED_FUTAI_TANI).Value
            newRow("futai_bunrui_nm") = sheet.Cells(index, COL_FUTAI_SELECTED_FUTAI_BUNRUI_NM).Value
            newRow("notax_flg") = sheet.Cells(index, COL_FUTAI_SELECTED_NOTAX_FLG_HIDDEN).Value
            newRow("shukei_grp") = sheet.Cells(index, COL_FUTAI_SELECTED_SHUKEI_GRP).Value
            newRow("kamoku_cd") = sheet.Cells(index, COL_FUTAI_SELECTED_KAMOKU_CD).Value
            newRow("saimoku_cd") = sheet.Cells(index, COL_FUTAI_SELECTED_SAIMOKU_CD).Value
            newRow("uchi_cd") = sheet.Cells(index, COL_FUTAI_SELECTED_UCHI_CD).Value
            newRow("shosai_cd") = sheet.Cells(index, COL_FUTAI_SELECTED_SHOSAI_CD).Value
            newRow("karikamoku_cd") = sheet.Cells(index, COL_FUTAI_SELECTED_KARIKAMOKU).Value
            newRow("kari_saimoku_cd") = sheet.Cells(index, COL_FUTAI_SELECTED_KARI_SAIMOKU_CD).Value
            newRow("kari_uchi_cd") = sheet.Cells(index, COL_FUTAI_SELECTED_KARI_UCHI_CD).Value
            newRow("kari_shosai_cd") = sheet.Cells(index, COL_FUTAI_SELECTED_KARI_SHOSAI_CD).Value
            newRow("kamoku_nm") = sheet.Cells(index, COL_FUTAI_SELECTED_KAMOKU_NM).Value
            newRow("saimoku_nm") = sheet.Cells(index, COL_FUTAI_SELECTED_SAIMOKU_NM).Value
            newRow("uchi_nm") = sheet.Cells(index, COL_FUTAI_SELECTED_UCHI_NM).Value
            newRow("shosai_nm") = sheet.Cells(index, COL_FUTAI_SELECTED_SHOSAI_NM).Value

            ' --- 2020/03/24 税区分追加対応 Start E.Okuda@Compass ---
            For i = 0 To UBound(arySpreadTitle)
                If arySpreadTitle(i) = sheet.Cells(index, COL_FUTAI_SELECTED_TAX_KBN).Value Then
                    newRow("tax_kbn") = i + 1
                End If
            Next

            '            newRow("tax_kbn") = sheet.Cells(index, COL_FUTAI_SELECTED_TAX_KBN).Value

            ' 税率
            newRow("tax_rate") = sheet.Cells(index, COL_FUTAI_SELECTED_ZEIRITSU).Value
            '            newRow("zeiritsu") = sheet.Cells(index, COL_FUTAI_SELECTED_ZEIRITSU).Value
            ' --- 2020/03/24 税区分追加対応 End E.Okuda@Compass ---

            ' 消費税額計算
            If sheet.Cells(index, COL_FUTAI_SELECTED_NOTAX_FLG_HIDDEN).Value = VAL_SOTOZEI Then
                intTaxkin = Long.Parse(sheet.Cells(index, COL_FUTAI_SELECTED_FUTAI_KIN).Value) * Long.Parse(sheet.Cells(index, COL_FUTAI_SELECTED_ZEIRITSU).Value) / 100
            Else
                intTaxkin = 0
            End If
            newRow("tax_kin") = intTaxkin

            table.Rows.Add(newRow)
            If isFirst = False Then
                futaiTotalNm = futaiTotalNm + "／" + sheet.Cells(index, COL_FUTAI_SELECTED_FUTAI_NM).Value
            Else
                futaiTotalNm = sheet.Cells(index, COL_FUTAI_SELECTED_FUTAI_NM).Value
                isFirst = False
            End If
            intTotalTaxkin += intTaxkin                 ' 2015.11.13 ADD h.hagiwara
            index = index + 1

            'newRow("yoyaku_dt") = dataEXTZ0209.PropStrRiyobi
            'newRow("futai_bunrui_cd") = sheet.Cells(index, 9).Value
            'newRow("futai_cd") = sheet.Cells(index, 10).Value
            'newRow("futai_tanka") = sheet.Cells(index, 2).Value
            'newRow("futai_su") = sheet.Cells(index, 3).Value
            'newRow("futai_shokei") = sheet.Cells(index, 5).Value
            'newRow("futai_chosei") = sheet.Cells(index, 6).Value
            'newRow("futai_kin") = sheet.Cells(index, 7).Value
            'newRow("futai_biko") = commonLogicEXT.DbNothingToNull(sheet.Cells(index, 8).Value)
            'newRow("futai_nm") = sheet.Cells(index, 1).Value
            'newRow("futai_tani") = sheet.Cells(index, 4).Value
            'newRow("futai_bunrui_nm") = sheet.Cells(index, 0).Value
            'newRow("notax_flg") = sheet.Cells(index, 11).Value
            'newRow("shukei_grp") = sheet.Cells(index, 12).Value
            'newRow("kamoku_cd") = sheet.Cells(index, 13).Value
            'newRow("saimoku_cd") = sheet.Cells(index, 14).Value
            'newRow("uchi_cd") = sheet.Cells(index, 15).Value
            'newRow("shosai_cd") = sheet.Cells(index, 16).Value
            'newRow("karikamoku_cd") = sheet.Cells(index, 17).Value
            'newRow("kari_saimoku_cd") = sheet.Cells(index, 18).Value
            'newRow("kari_uchi_cd") = sheet.Cells(index, 19).Value
            'newRow("kari_shosai_cd") = sheet.Cells(index, 20).Value
            'newRow("kamoku_nm") = sheet.Cells(index, 21).Value
            'newRow("saimoku_nm") = sheet.Cells(index, 22).Value
            'newRow("uchi_nm") = sheet.Cells(index, 23).Value
            'newRow("shosai_nm") = sheet.Cells(index, 24).Value
            '' 2015.11.13 ADD START↓ h.hagiwara
            'If sheet.Cells(index, 11).Value = "0" Then
            '    'intTaxkin = Integer.Parse(sheet.Cells(index, 7).Value) * dblTaxritu              ' 2015.12.21 UPD h.hagiwara
            '    intTaxkin = Long.Parse(sheet.Cells(index, 7).Value) * dblTaxritu              ' 2015.12.21 UPD h.hagiwara
            'Else
            '    intTaxkin = 0
            'End If
            'newRow("tax_kin") = intTaxkin
            '' 2015.11.13 ADD END↑ h.hagiwara
            'table.Rows.Add(newRow)
            'If isFirst = False Then
            '    futaiTotalNm = futaiTotalNm + "／" + sheet.Cells(index, 1).Value
            'Else
            '    futaiTotalNm = sheet.Cells(index, 1).Value
            '    isFirst = False
            'End If
            'intTotalTaxkin += intTaxkin                 ' 2015.11.13 ADD h.hagiwara
            'index = index + 1
        Loop
        If String.IsNullOrEmpty(Me.lblTotal.Text) = True Then
            dataEXTZ0209.PropIntTotal = 0
            dataEXTZ0209.PropIntTotalTax = 0                                                       ' 2015.11.13 ADD h.hagiwara
        Else
            'dataEXTZ0209.PropIntTotal = Integer.Parse(Me.lblTotal.Text.Replace(",", ""))           ' 2015.12.21 UPD h.hagiwara
            dataEXTZ0209.PropIntTotal = Long.Parse(Me.lblTotal.Text.Replace(",", ""))               ' 2015.12.21 UPD h.hagiwara
            dataEXTZ0209.PropIntTotalTax = intTotalTaxkin                                           ' 2015.11.13 ADD h.hagiwara
        End If
        dataEXTZ0209.PropStrFutaiTotalNm = futaiTotalNm
        dataEXTZ0209.PropFutaiDetailTable = table
        ppBlnChangeFlg = True
        Me.Close()
    End Sub

    ''' <summary>
    ''' 戻る処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnBack_Click(sender As Object, e As EventArgs) Handles btnBack.Click
        ppBlnChangeFlg = False
        Me.Close()
    End Sub

    ''' <summary>
    ''' プロパティセット【変更フラグ】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDsApproval</returns>
    ''' <remarks><para>作成情報：2015/08/25 k.machida
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

    ' --- 2019/06/10 軽減税率対応 Start E.Okuda@Compass ---

    ''' <summary>
    ''' 合計(税込)再計算
    ''' </summary>
    ''' <param name="objSheet">選択スプレッドシート</param>
    ''' <param name="intRowNo">行番号</param>
    ''' <remarks> 選択スプレッドシートの合計(税込)を再計算する。
    ''' <para>作成情報：2019/06/10 E.Okuda@Compass
    ''' <p>改訂情報：</p>
    ''' </para>
    ''' </remarks>
    Private Sub RecalcAmountTaxIncluded(ByRef objSheet As FarPoint.Win.Spread.SheetView, ByVal intRowNo As Integer)
        Dim lngChosei As Long

        ' シートの有効行数をデータテーブルの行数にする。
        '        objSheet.RowCount = dtbFutaiDetail.Rows.Count

        ' 数量、調整額をスプレッドシートから取得する。
        Dim strQuantity As String = objSheet.Cells(intRowNo, COL_FUTAI_SELECTED_FUTAI_SU).Value
        Dim strChosei As String = objSheet.Cells(intRowNo, COL_FUTAI_SELECTED_FUTAI_CHOSEI).Value
        Dim strTotal As String = objSheet.Cells(intRowNo, COL_FUTAI_SELECTED_FUTAI_KIN).Value

        ' 取得した値をチェックする。
        If commonValidate.IsHalfNmb(strQuantity) = False Then
            MsgBox(String.Format(CommonDeclareEXT.E0003, "数量"), MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
        If commonValidate.IsHalfMNmb(strChosei) = False Then
            MsgBox(String.Format(CommonDeclareEXT.E0004, "調整金額"), MsgBoxStyle.Exclamation, "エラー")
            Exit Sub
        End If
        If String.IsNullOrEmpty(strQuantity) = True Then
            objSheet.Cells(intRowNo, COL_FUTAI_SELECTED_FUTAI_SU).Value = "1"
        End If
        If String.IsNullOrEmpty(strChosei) = True Then
            objSheet.Cells(intRowNo, COL_FUTAI_SELECTED_FUTAI_CHOSEI).Value = "0"
            lngChosei = 0
        End If

        ' 税区分と税率を取得して合計(税込)を算出
        If objSheet.Cells(intRowNo, COL_FUTAI_SELECTED_NOTAX_FLG).Value = DISP_SOTOZEI Then
            objSheet.Cells(intRowNo, COL_FUTAI_SELECTED_FUTAI_KIN_TAX).Value =
            Long.Parse(strTotal) + Long.Parse(strTotal) * Integer.Parse(objSheet.Cells(intRowNo, COL_FUTAI_SELECTED_ZEIRITSU).Value) / 100
        ElseIf objSheet.Cells(intRowNo, COL_FUTAI_SELECTED_NOTAX_FLG).Value = VAL_UCHIZEI Then
            objSheet.Cells(intRowNo, COL_FUTAI_SELECTED_FUTAI_KIN_TAX).Value = Long.Parse(strTotal)
        End If
    End Sub

    ' --- 2019/06/10 軽減税率対応 End E.Okuda@Compass ---

End Class
