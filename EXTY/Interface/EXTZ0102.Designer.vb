<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class EXTZ0102
    Inherits System.Windows.Forms.Form

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows フォーム デザイナーで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャは Windows フォーム デザイナーで必要です。
    'Windows フォーム デザイナーを使用して変更できます。  
    'コード エディターを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim ButtonCellType1 As FarPoint.Win.Spread.CellType.ButtonCellType = New FarPoint.Win.Spread.CellType.ButtonCellType()
        Dim ButtonCellType2 As FarPoint.Win.Spread.CellType.ButtonCellType = New FarPoint.Win.Spread.CellType.ButtonCellType()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(EXTZ0102))
        Me.Label65 = New System.Windows.Forms.Label()
        Me.Label64 = New System.Windows.Forms.Label()
        Me.txtAitesakiNm_Kana = New System.Windows.Forms.TextBox()
        Me.Label63 = New System.Windows.Forms.Label()
        Me.txtSaijiShutsuenNm = New System.Windows.Forms.TextBox()
        Me.Label62 = New System.Windows.Forms.Label()
        Me.txtSeikyuIraiNo = New System.Windows.Forms.TextBox()
        Me.Label54 = New System.Windows.Forms.Label()
        Me.Label53 = New System.Windows.Forms.Label()
        Me.txtRiyoNm = New System.Windows.Forms.TextBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.rdoStudio = New System.Windows.Forms.RadioButton()
        Me.rdoTheatre = New System.Windows.Forms.RadioButton()
        Me.chkMinyukinOnly = New System.Windows.Forms.CheckBox()
        Me.btnAitesaakiSearch = New System.Windows.Forms.Button()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.dtpNyukinDt_To = New Common.DateTimePickerEx()
        Me.dtpNyukinDt_From = New Common.DateTimePickerEx()
        Me.dtpNyukinYoteiDt_To = New Common.DateTimePickerEx()
        Me.dtpNyukinYoteiDt_From = New Common.DateTimePickerEx()
        Me.dtpSeikyuDt_From = New Common.DateTimePickerEx()
        Me.dtpSeikyuDt_To = New Common.DateTimePickerEx()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtAitesakiNm = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnSearch = New System.Windows.Forms.Button()
        Me.Label66 = New System.Windows.Forms.Label()
        Me.btnBack = New System.Windows.Forms.Button()
        Me.vwSeikyuTheatre = New FarPoint.Win.Spread.FpSpread()
        Me.vwSeikyuTheatre_Sheet1 = New FarPoint.Win.Spread.SheetView()
        Me.vwSeikyuStudio = New FarPoint.Win.Spread.FpSpread()
        Me.vwSeikyuStudio_Sheet1 = New FarPoint.Win.Spread.SheetView()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Panel1.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        CType(Me.vwSeikyuTheatre, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.vwSeikyuTheatre_Sheet1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.vwSeikyuStudio, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.vwSeikyuStudio_Sheet1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label65
        '
        Me.Label65.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label65.Location = New System.Drawing.Point(21, 66)
        Me.Label65.Name = "Label65"
        Me.Label65.Size = New System.Drawing.Size(85, 21)
        Me.Label65.TabIndex = 2
        Me.Label65.Text = "請求日"
        '
        'Label64
        '
        Me.Label64.AutoSize = True
        Me.Label64.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label64.Location = New System.Drawing.Point(21, 99)
        Me.Label64.Name = "Label64"
        Me.Label64.Size = New System.Drawing.Size(82, 13)
        Me.Label64.TabIndex = 14
        Me.Label64.Text = "相手先名(ｶﾅ)" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'txtAitesakiNm_Kana
        '
        Me.txtAitesakiNm_Kana.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txtAitesakiNm_Kana.Location = New System.Drawing.Point(224, 96)
        Me.txtAitesakiNm_Kana.Name = "txtAitesakiNm_Kana"
        Me.txtAitesakiNm_Kana.Size = New System.Drawing.Size(306, 20)
        Me.txtAitesakiNm_Kana.TabIndex = 15
        '
        'Label63
        '
        Me.Label63.AutoSize = True
        Me.Label63.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label63.Location = New System.Drawing.Point(646, 134)
        Me.Label63.Name = "Label63"
        Me.Label63.Size = New System.Drawing.Size(129, 13)
        Me.Label63.TabIndex = 21
        Me.Label63.Text = "催事名／アーティスト名"
        '
        'txtSaijiShutsuenNm
        '
        Me.txtSaijiShutsuenNm.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txtSaijiShutsuenNm.Location = New System.Drawing.Point(809, 131)
        Me.txtSaijiShutsuenNm.Name = "txtSaijiShutsuenNm"
        Me.txtSaijiShutsuenNm.Size = New System.Drawing.Size(306, 20)
        Me.txtSaijiShutsuenNm.TabIndex = 22
        '
        'Label62
        '
        Me.Label62.AutoSize = True
        Me.Label62.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label62.Location = New System.Drawing.Point(21, 171)
        Me.Label62.Name = "Label62"
        Me.Label62.Size = New System.Drawing.Size(85, 13)
        Me.Label62.TabIndex = 23
        Me.Label62.Text = "請求依頼番号"
        '
        'txtSeikyuIraiNo
        '
        Me.txtSeikyuIraiNo.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txtSeikyuIraiNo.Location = New System.Drawing.Point(224, 168)
        Me.txtSeikyuIraiNo.Name = "txtSeikyuIraiNo"
        Me.txtSeikyuIraiNo.Size = New System.Drawing.Size(112, 20)
        Me.txtSeikyuIraiNo.TabIndex = 24
        '
        'Label54
        '
        Me.Label54.AutoSize = True
        Me.Label54.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label54.Location = New System.Drawing.Point(21, 134)
        Me.Label54.Name = "Label54"
        Me.Label54.Size = New System.Drawing.Size(59, 13)
        Me.Label54.TabIndex = 19
        Me.Label54.Text = "利用者名"
        '
        'Label53
        '
        Me.Label53.AutoSize = True
        Me.Label53.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label53.Location = New System.Drawing.Point(316, 66)
        Me.Label53.Name = "Label53"
        Me.Label53.Size = New System.Drawing.Size(20, 13)
        Me.Label53.TabIndex = 4
        Me.Label53.Text = "～"
        '
        'txtRiyoNm
        '
        Me.txtRiyoNm.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txtRiyoNm.Location = New System.Drawing.Point(224, 131)
        Me.txtRiyoNm.Name = "txtRiyoNm"
        Me.txtRiyoNm.Size = New System.Drawing.Size(306, 20)
        Me.txtRiyoNm.TabIndex = 20
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.rdoStudio)
        Me.Panel1.Controls.Add(Me.rdoTheatre)
        Me.Panel1.Location = New System.Drawing.Point(153, 30)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(222, 22)
        Me.Panel1.TabIndex = 1
        '
        'rdoStudio
        '
        Me.rdoStudio.AutoSize = True
        Me.rdoStudio.Location = New System.Drawing.Point(99, 3)
        Me.rdoStudio.Name = "rdoStudio"
        Me.rdoStudio.Size = New System.Drawing.Size(65, 17)
        Me.rdoStudio.TabIndex = 1
        Me.rdoStudio.Text = "スタジオ"
        Me.rdoStudio.UseVisualStyleBackColor = True
        '
        'rdoTheatre
        '
        Me.rdoTheatre.AutoSize = True
        Me.rdoTheatre.Checked = True
        Me.rdoTheatre.Location = New System.Drawing.Point(3, 3)
        Me.rdoTheatre.Name = "rdoTheatre"
        Me.rdoTheatre.Size = New System.Drawing.Size(66, 17)
        Me.rdoTheatre.TabIndex = 0
        Me.rdoTheatre.TabStop = True
        Me.rdoTheatre.Text = "シアター"
        Me.rdoTheatre.UseVisualStyleBackColor = True
        '
        'chkMinyukinOnly
        '
        Me.chkMinyukinOnly.AutoSize = True
        Me.chkMinyukinOnly.Checked = True
        Me.chkMinyukinOnly.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkMinyukinOnly.Location = New System.Drawing.Point(24, 204)
        Me.chkMinyukinOnly.Name = "chkMinyukinOnly"
        Me.chkMinyukinOnly.Size = New System.Drawing.Size(248, 17)
        Me.chkMinyukinOnly.TabIndex = 25
        Me.chkMinyukinOnly.Text = "入金予定日を過ぎた未入金請求のみ表示"
        Me.chkMinyukinOnly.UseVisualStyleBackColor = True
        '
        'btnAitesaakiSearch
        '
        Me.btnAitesaakiSearch.BackgroundImage = Global.EXTY.My.Resources.Resources.マスタ検索
        Me.btnAitesaakiSearch.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnAitesaakiSearch.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.btnAitesaakiSearch.Location = New System.Drawing.Point(536, 92)
        Me.btnAitesaakiSearch.Name = "btnAitesaakiSearch"
        Me.btnAitesaakiSearch.Size = New System.Drawing.Size(36, 27)
        Me.btnAitesaakiSearch.TabIndex = 16
        Me.btnAitesaakiSearch.UseVisualStyleBackColor = True
        '
        'GroupBox5
        '
        Me.GroupBox5.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox5.Controls.Add(Me.dtpNyukinDt_To)
        Me.GroupBox5.Controls.Add(Me.dtpNyukinDt_From)
        Me.GroupBox5.Controls.Add(Me.dtpNyukinYoteiDt_To)
        Me.GroupBox5.Controls.Add(Me.dtpNyukinYoteiDt_From)
        Me.GroupBox5.Controls.Add(Me.dtpSeikyuDt_From)
        Me.GroupBox5.Controls.Add(Me.dtpSeikyuDt_To)
        Me.GroupBox5.Controls.Add(Me.Label5)
        Me.GroupBox5.Controls.Add(Me.txtAitesakiNm)
        Me.GroupBox5.Controls.Add(Me.Label3)
        Me.GroupBox5.Controls.Add(Me.Label4)
        Me.GroupBox5.Controls.Add(Me.Label2)
        Me.GroupBox5.Controls.Add(Me.Label1)
        Me.GroupBox5.Controls.Add(Me.btnSearch)
        Me.GroupBox5.Controls.Add(Me.btnAitesaakiSearch)
        Me.GroupBox5.Controls.Add(Me.chkMinyukinOnly)
        Me.GroupBox5.Controls.Add(Me.Panel1)
        Me.GroupBox5.Controls.Add(Me.txtRiyoNm)
        Me.GroupBox5.Controls.Add(Me.Label53)
        Me.GroupBox5.Controls.Add(Me.Label54)
        Me.GroupBox5.Controls.Add(Me.txtSeikyuIraiNo)
        Me.GroupBox5.Controls.Add(Me.Label62)
        Me.GroupBox5.Controls.Add(Me.txtSaijiShutsuenNm)
        Me.GroupBox5.Controls.Add(Me.Label63)
        Me.GroupBox5.Controls.Add(Me.txtAitesakiNm_Kana)
        Me.GroupBox5.Controls.Add(Me.Label64)
        Me.GroupBox5.Controls.Add(Me.Label65)
        Me.GroupBox5.Controls.Add(Me.Label66)
        Me.GroupBox5.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.GroupBox5.Location = New System.Drawing.Point(12, 42)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(1539, 252)
        Me.GroupBox5.TabIndex = 0
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "検索条件"
        '
        'dtpNyukinDt_To
        '
        Me.dtpNyukinDt_To.CausesValidation = False
        Me.dtpNyukinDt_To.Location = New System.Drawing.Point(1251, 57)
        Me.dtpNyukinDt_To.Name = "dtpNyukinDt_To"
        Me.dtpNyukinDt_To.Size = New System.Drawing.Size(150, 31)
        Me.dtpNyukinDt_To.TabIndex = 13
        '
        'dtpNyukinDt_From
        '
        Me.dtpNyukinDt_From.CausesValidation = False
        Me.dtpNyukinDt_From.Location = New System.Drawing.Point(1059, 57)
        Me.dtpNyukinDt_From.Name = "dtpNyukinDt_From"
        Me.dtpNyukinDt_From.Size = New System.Drawing.Size(150, 29)
        Me.dtpNyukinDt_From.TabIndex = 11
        '
        'dtpNyukinYoteiDt_To
        '
        Me.dtpNyukinYoteiDt_To.CausesValidation = False
        Me.dtpNyukinYoteiDt_To.Location = New System.Drawing.Point(812, 58)
        Me.dtpNyukinYoteiDt_To.Name = "dtpNyukinYoteiDt_To"
        Me.dtpNyukinYoteiDt_To.Size = New System.Drawing.Size(158, 29)
        Me.dtpNyukinYoteiDt_To.TabIndex = 9
        '
        'dtpNyukinYoteiDt_From
        '
        Me.dtpNyukinYoteiDt_From.CausesValidation = False
        Me.dtpNyukinYoteiDt_From.Location = New System.Drawing.Point(617, 59)
        Me.dtpNyukinYoteiDt_From.Name = "dtpNyukinYoteiDt_From"
        Me.dtpNyukinYoteiDt_From.Size = New System.Drawing.Size(147, 27)
        Me.dtpNyukinYoteiDt_From.TabIndex = 7
        '
        'dtpSeikyuDt_From
        '
        Me.dtpSeikyuDt_From.CausesValidation = False
        Me.dtpSeikyuDt_From.Location = New System.Drawing.Point(153, 61)
        Me.dtpSeikyuDt_From.Name = "dtpSeikyuDt_From"
        Me.dtpSeikyuDt_From.Size = New System.Drawing.Size(157, 33)
        Me.dtpSeikyuDt_From.TabIndex = 3
        '
        'dtpSeikyuDt_To
        '
        Me.dtpSeikyuDt_To.Location = New System.Drawing.Point(352, 60)
        Me.dtpSeikyuDt_To.Name = "dtpSeikyuDt_To"
        Me.dtpSeikyuDt_To.Size = New System.Drawing.Size(153, 30)
        Me.dtpSeikyuDt_To.TabIndex = 5
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label5.Location = New System.Drawing.Point(646, 99)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(59, 13)
        Me.Label5.TabIndex = 17
        Me.Label5.Text = "相手先名"
        '
        'txtAitesakiNm
        '
        Me.txtAitesakiNm.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txtAitesakiNm.Location = New System.Drawing.Point(809, 96)
        Me.txtAitesakiNm.Name = "txtAitesakiNm"
        Me.txtAitesakiNm.Size = New System.Drawing.Size(306, 20)
        Me.txtAitesakiNm.TabIndex = 18
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label3.Location = New System.Drawing.Point(1216, 61)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(20, 13)
        Me.Label3.TabIndex = 12
        Me.Label3.Text = "～"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label4.Location = New System.Drawing.Point(1007, 62)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(46, 13)
        Me.Label4.TabIndex = 10
        Me.Label4.Text = "入金日"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label2.Location = New System.Drawing.Point(777, 64)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(20, 13)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "～"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label1.Location = New System.Drawing.Point(538, 63)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(72, 13)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "入金予定日"
        '
        'btnSearch
        '
        Me.btnSearch.Location = New System.Drawing.Point(1349, 216)
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.Size = New System.Drawing.Size(96, 23)
        Me.btnSearch.TabIndex = 26
        Me.btnSearch.Text = "検索"
        Me.btnSearch.UseVisualStyleBackColor = True
        '
        'Label66
        '
        Me.Label66.AutoSize = True
        Me.Label66.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label66.Location = New System.Drawing.Point(21, 34)
        Me.Label66.Name = "Label66"
        Me.Label66.Size = New System.Drawing.Size(59, 13)
        Me.Label66.TabIndex = 0
        Me.Label66.Text = "検索対象"
        '
        'btnBack
        '
        Me.btnBack.Location = New System.Drawing.Point(758, 904)
        Me.btnBack.Name = "btnBack"
        Me.btnBack.Size = New System.Drawing.Size(175, 36)
        Me.btnBack.TabIndex = 3
        Me.btnBack.Text = "戻る"
        Me.btnBack.UseVisualStyleBackColor = True
        '
        'vwSeikyuTheatre
        '
        Me.vwSeikyuTheatre.AccessibleDescription = "vwSeikyuTheatre, Sheet1, Row 0, Column 0, "
        Me.vwSeikyuTheatre.Location = New System.Drawing.Point(12, 310)
        Me.vwSeikyuTheatre.Name = "vwSeikyuTheatre"
        Me.vwSeikyuTheatre.Sheets.AddRange(New FarPoint.Win.Spread.SheetView() {Me.vwSeikyuTheatre_Sheet1})
        Me.vwSeikyuTheatre.Size = New System.Drawing.Size(1552, 563)
        Me.vwSeikyuTheatre.TabIndex = 1
        '
        'vwSeikyuTheatre_Sheet1
        '
        Me.vwSeikyuTheatre_Sheet1.Reset()
        Me.vwSeikyuTheatre_Sheet1.SheetName = "Sheet1"
        'Formulas and custom names must be loaded with R1C1 reference style
        Me.vwSeikyuTheatre_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.R1C1
        Me.vwSeikyuTheatre_Sheet1.ColumnCount = 14
        Me.vwSeikyuTheatre_Sheet1.Cells.Get(0, 0).Locked = True
        Me.vwSeikyuTheatre_Sheet1.ColumnHeader.Cells.Get(0, 0).Value = "請求依頼番号"
        Me.vwSeikyuTheatre_Sheet1.ColumnHeader.Cells.Get(0, 1).Value = "請求日"
        Me.vwSeikyuTheatre_Sheet1.ColumnHeader.Cells.Get(0, 2).Value = "利用開始日"
        Me.vwSeikyuTheatre_Sheet1.ColumnHeader.Cells.Get(0, 3).Value = "利用終了日"
        Me.vwSeikyuTheatre_Sheet1.ColumnHeader.Cells.Get(0, 4).Value = "催事名"
        Me.vwSeikyuTheatre_Sheet1.ColumnHeader.Cells.Get(0, 5).Value = "相手先名"
        Me.vwSeikyuTheatre_Sheet1.ColumnHeader.Cells.Get(0, 6).Value = "請求金額"
        Me.vwSeikyuTheatre_Sheet1.ColumnHeader.Cells.Get(0, 7).Value = "請求内容"
        Me.vwSeikyuTheatre_Sheet1.ColumnHeader.Cells.Get(0, 8).Value = "入金予定日"
        Me.vwSeikyuTheatre_Sheet1.ColumnHeader.Cells.Get(0, 9).Value = "入金日"
        Me.vwSeikyuTheatre_Sheet1.ColumnHeader.Cells.Get(0, 10).Value = "入金額"
        Me.vwSeikyuTheatre_Sheet1.ColumnHeader.Cells.Get(0, 11).Value = "入金状況"
        Me.vwSeikyuTheatre_Sheet1.ColumnHeader.Cells.Get(0, 12).Value = "予約番号"
        Me.vwSeikyuTheatre_Sheet1.ColumnHeader.Cells.Get(0, 13).Value = "　"
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(0).AllowAutoSort = True
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(0).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(0).Label = "請求依頼番号"
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(0).Locked = True
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(0).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(0).Width = 104.0!
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(1).AllowAutoSort = True
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(1).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(1).Label = "請求日"
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(1).Locked = True
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(1).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(1).Width = 103.0!
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(2).AllowAutoSort = True
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(2).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(2).Label = "利用開始日"
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(2).Locked = True
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(2).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(2).Width = 103.0!
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(3).AllowAutoSort = True
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(3).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(3).Label = "利用終了日"
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(3).Locked = True
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(3).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(3).Width = 103.0!
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(4).Label = "催事名"
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(4).Locked = True
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(4).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(4).Width = 261.0!
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(5).AllowAutoSort = True
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(5).Label = "相手先名"
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(5).Locked = True
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(5).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(5).Width = 193.0!
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(6).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Right
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(6).Label = "請求金額"
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(6).Locked = True
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(6).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(6).Width = 104.0!
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(7).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Left
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(7).Label = "請求内容"
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(7).Locked = True
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(7).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(7).Width = 158.0!
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(8).AllowAutoSort = True
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(8).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(8).Label = "入金予定日"
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(8).Locked = True
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(8).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(8).Width = 103.0!
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(9).AllowAutoSort = True
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(9).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(9).Label = "入金日"
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(9).Locked = True
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(9).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(9).Width = 103.0!
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(10).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Right
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(10).Label = "入金額"
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(10).Locked = True
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(10).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(10).Width = 104.0!
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(11).AllowAutoSort = True
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(11).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(11).Label = "入金状況"
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(11).Locked = True
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(11).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(11).Width = 80.0!
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(12).AllowAutoSort = True
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(12).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(12).Label = "予約番号"
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(12).Locked = True
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(12).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(12).Width = 90.0!
        ButtonCellType1.ButtonColor2 = System.Drawing.SystemColors.ButtonFace
        ButtonCellType1.Text = "確認"
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(13).CellType = ButtonCellType1
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(13).Label = "　"
        Me.vwSeikyuTheatre_Sheet1.Columns.Get(13).Width = 70.0!
        Me.vwSeikyuTheatre_Sheet1.RowHeader.Columns.Default.Resizable = False
        Me.vwSeikyuTheatre_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.A1
        '
        'vwSeikyuStudio
        '
        Me.vwSeikyuStudio.AccessibleDescription = "vwSeikyuStudio, Sheet1, Row 0, Column 0, "
        Me.vwSeikyuStudio.Location = New System.Drawing.Point(12, 310)
        Me.vwSeikyuStudio.Name = "vwSeikyuStudio"
        Me.vwSeikyuStudio.Sheets.AddRange(New FarPoint.Win.Spread.SheetView() {Me.vwSeikyuStudio_Sheet1})
        Me.vwSeikyuStudio.Size = New System.Drawing.Size(1640, 563)
        Me.vwSeikyuStudio.TabIndex = 2
        '
        'vwSeikyuStudio_Sheet1
        '
        Me.vwSeikyuStudio_Sheet1.Reset()
        Me.vwSeikyuStudio_Sheet1.SheetName = "Sheet1"
        'Formulas and custom names must be loaded with R1C1 reference style
        Me.vwSeikyuStudio_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.R1C1
        Me.vwSeikyuStudio_Sheet1.ColumnCount = 15
        Me.vwSeikyuStudio_Sheet1.ColumnHeader.Cells.Get(0, 0).Value = "請求依頼番号"
        Me.vwSeikyuStudio_Sheet1.ColumnHeader.Cells.Get(0, 1).Value = "請求日"
        Me.vwSeikyuStudio_Sheet1.ColumnHeader.Cells.Get(0, 2).Value = "利用開始日"
        Me.vwSeikyuStudio_Sheet1.ColumnHeader.Cells.Get(0, 3).Value = "利用終了日"
        Me.vwSeikyuStudio_Sheet1.ColumnHeader.Cells.Get(0, 4).Value = "スタジオ"
        Me.vwSeikyuStudio_Sheet1.ColumnHeader.Cells.Get(0, 5).Value = "アーティスト名"
        Me.vwSeikyuStudio_Sheet1.ColumnHeader.Cells.Get(0, 6).Value = "相手先名"
        Me.vwSeikyuStudio_Sheet1.ColumnHeader.Cells.Get(0, 7).Value = "請求金額"
        Me.vwSeikyuStudio_Sheet1.ColumnHeader.Cells.Get(0, 8).Value = "請求内容"
        Me.vwSeikyuStudio_Sheet1.ColumnHeader.Cells.Get(0, 9).Value = "入金予定日"
        Me.vwSeikyuStudio_Sheet1.ColumnHeader.Cells.Get(0, 10).Value = "入金日"
        Me.vwSeikyuStudio_Sheet1.ColumnHeader.Cells.Get(0, 11).Value = "入金額"
        Me.vwSeikyuStudio_Sheet1.ColumnHeader.Cells.Get(0, 12).Value = "入金状況"
        Me.vwSeikyuStudio_Sheet1.ColumnHeader.Cells.Get(0, 13).Value = "予約番号"
        Me.vwSeikyuStudio_Sheet1.ColumnHeader.Cells.Get(0, 14).Value = "　"
        Me.vwSeikyuStudio_Sheet1.Columns.Get(0).AllowAutoSort = True
        Me.vwSeikyuStudio_Sheet1.Columns.Get(0).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwSeikyuStudio_Sheet1.Columns.Get(0).Label = "請求依頼番号"
        Me.vwSeikyuStudio_Sheet1.Columns.Get(0).Locked = True
        Me.vwSeikyuStudio_Sheet1.Columns.Get(0).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwSeikyuStudio_Sheet1.Columns.Get(0).Width = 104.0!
        Me.vwSeikyuStudio_Sheet1.Columns.Get(1).AllowAutoSort = True
        Me.vwSeikyuStudio_Sheet1.Columns.Get(1).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwSeikyuStudio_Sheet1.Columns.Get(1).Label = "請求日"
        Me.vwSeikyuStudio_Sheet1.Columns.Get(1).Locked = True
        Me.vwSeikyuStudio_Sheet1.Columns.Get(1).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwSeikyuStudio_Sheet1.Columns.Get(1).Width = 103.0!
        Me.vwSeikyuStudio_Sheet1.Columns.Get(2).AllowAutoSort = True
        Me.vwSeikyuStudio_Sheet1.Columns.Get(2).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwSeikyuStudio_Sheet1.Columns.Get(2).Label = "利用開始日"
        Me.vwSeikyuStudio_Sheet1.Columns.Get(2).Locked = True
        Me.vwSeikyuStudio_Sheet1.Columns.Get(2).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwSeikyuStudio_Sheet1.Columns.Get(2).Width = 103.0!
        Me.vwSeikyuStudio_Sheet1.Columns.Get(3).AllowAutoSort = True
        Me.vwSeikyuStudio_Sheet1.Columns.Get(3).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwSeikyuStudio_Sheet1.Columns.Get(3).Label = "利用終了日"
        Me.vwSeikyuStudio_Sheet1.Columns.Get(3).Locked = True
        Me.vwSeikyuStudio_Sheet1.Columns.Get(3).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwSeikyuStudio_Sheet1.Columns.Get(3).Width = 103.0!
        Me.vwSeikyuStudio_Sheet1.Columns.Get(4).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwSeikyuStudio_Sheet1.Columns.Get(4).Label = "スタジオ"
        Me.vwSeikyuStudio_Sheet1.Columns.Get(4).Locked = True
        Me.vwSeikyuStudio_Sheet1.Columns.Get(4).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwSeikyuStudio_Sheet1.Columns.Get(4).Width = 88.0!
        Me.vwSeikyuStudio_Sheet1.Columns.Get(5).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Left
        Me.vwSeikyuStudio_Sheet1.Columns.Get(5).Label = "アーティスト名"
        Me.vwSeikyuStudio_Sheet1.Columns.Get(5).Locked = True
        Me.vwSeikyuStudio_Sheet1.Columns.Get(5).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwSeikyuStudio_Sheet1.Columns.Get(5).Width = 261.0!
        Me.vwSeikyuStudio_Sheet1.Columns.Get(6).AllowAutoSort = True
        Me.vwSeikyuStudio_Sheet1.Columns.Get(6).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Left
        Me.vwSeikyuStudio_Sheet1.Columns.Get(6).Label = "相手先名"
        Me.vwSeikyuStudio_Sheet1.Columns.Get(6).Locked = True
        Me.vwSeikyuStudio_Sheet1.Columns.Get(6).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwSeikyuStudio_Sheet1.Columns.Get(6).Width = 193.0!
        Me.vwSeikyuStudio_Sheet1.Columns.Get(7).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Right
        Me.vwSeikyuStudio_Sheet1.Columns.Get(7).Label = "請求金額"
        Me.vwSeikyuStudio_Sheet1.Columns.Get(7).Locked = True
        Me.vwSeikyuStudio_Sheet1.Columns.Get(7).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwSeikyuStudio_Sheet1.Columns.Get(7).Width = 104.0!
        Me.vwSeikyuStudio_Sheet1.Columns.Get(8).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Left
        Me.vwSeikyuStudio_Sheet1.Columns.Get(8).Label = "請求内容"
        Me.vwSeikyuStudio_Sheet1.Columns.Get(8).Locked = True
        Me.vwSeikyuStudio_Sheet1.Columns.Get(8).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwSeikyuStudio_Sheet1.Columns.Get(8).Width = 158.0!
        Me.vwSeikyuStudio_Sheet1.Columns.Get(9).AllowAutoSort = True
        Me.vwSeikyuStudio_Sheet1.Columns.Get(9).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwSeikyuStudio_Sheet1.Columns.Get(9).Label = "入金予定日"
        Me.vwSeikyuStudio_Sheet1.Columns.Get(9).Locked = True
        Me.vwSeikyuStudio_Sheet1.Columns.Get(9).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwSeikyuStudio_Sheet1.Columns.Get(9).Width = 103.0!
        Me.vwSeikyuStudio_Sheet1.Columns.Get(10).AllowAutoSort = True
        Me.vwSeikyuStudio_Sheet1.Columns.Get(10).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwSeikyuStudio_Sheet1.Columns.Get(10).Label = "入金日"
        Me.vwSeikyuStudio_Sheet1.Columns.Get(10).Locked = True
        Me.vwSeikyuStudio_Sheet1.Columns.Get(10).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwSeikyuStudio_Sheet1.Columns.Get(10).Width = 103.0!
        Me.vwSeikyuStudio_Sheet1.Columns.Get(11).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Right
        Me.vwSeikyuStudio_Sheet1.Columns.Get(11).Label = "入金額"
        Me.vwSeikyuStudio_Sheet1.Columns.Get(11).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwSeikyuStudio_Sheet1.Columns.Get(11).Width = 104.0!
        Me.vwSeikyuStudio_Sheet1.Columns.Get(12).AllowAutoSort = True
        Me.vwSeikyuStudio_Sheet1.Columns.Get(12).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwSeikyuStudio_Sheet1.Columns.Get(12).Label = "入金状況"
        Me.vwSeikyuStudio_Sheet1.Columns.Get(12).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwSeikyuStudio_Sheet1.Columns.Get(12).Width = 80.0!
        Me.vwSeikyuStudio_Sheet1.Columns.Get(13).AllowAutoSort = True
        Me.vwSeikyuStudio_Sheet1.Columns.Get(13).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwSeikyuStudio_Sheet1.Columns.Get(13).Label = "予約番号"
        Me.vwSeikyuStudio_Sheet1.Columns.Get(13).Locked = True
        Me.vwSeikyuStudio_Sheet1.Columns.Get(13).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwSeikyuStudio_Sheet1.Columns.Get(13).Width = 90.0!
        ButtonCellType2.ButtonColor2 = System.Drawing.SystemColors.ButtonFace
        ButtonCellType2.Text = "確認"
        Me.vwSeikyuStudio_Sheet1.Columns.Get(14).CellType = ButtonCellType2
        Me.vwSeikyuStudio_Sheet1.Columns.Get(14).Label = "　"
        Me.vwSeikyuStudio_Sheet1.Columns.Get(14).Width = 70.0!
        Me.vwSeikyuStudio_Sheet1.RowHeader.Columns.Default.Resizable = False
        Me.vwSeikyuStudio_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.A1
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label6.Location = New System.Drawing.Point(20, 880)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(494, 13)
        Me.Label6.TabIndex = 24
        Me.Label6.Text = "※表のデータは、コピーしたい部分を選択し、キーボードの「Ctrl」＋「C」を押すとコピーできます。"
        '
        'EXTZ0102
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.BackColor = System.Drawing.Color.White
        Me.BackgroundImage = Global.EXTY.My.Resources.Resources.背景
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(1660, 962)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.GroupBox5)
        Me.Controls.Add(Me.btnBack)
        Me.Controls.Add(Me.vwSeikyuTheatre)
        Me.Controls.Add(Me.vwSeikyuStudio)
        Me.DoubleBuffered = True
        Me.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "EXTZ0102"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ちゃり～ん。　請求一覧(共通)"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox5.PerformLayout()
        CType(Me.vwSeikyuTheatre, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.vwSeikyuTheatre_Sheet1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.vwSeikyuStudio, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.vwSeikyuStudio_Sheet1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label65 As System.Windows.Forms.Label
    Friend WithEvents Label64 As System.Windows.Forms.Label
    Friend WithEvents txtAitesakiNm_Kana As System.Windows.Forms.TextBox
    Friend WithEvents Label63 As System.Windows.Forms.Label
    Friend WithEvents txtSaijiShutsuenNm As System.Windows.Forms.TextBox
    Friend WithEvents Label62 As System.Windows.Forms.Label
    Friend WithEvents txtSeikyuIraiNo As System.Windows.Forms.TextBox
    Friend WithEvents Label54 As System.Windows.Forms.Label
    Friend WithEvents Label53 As System.Windows.Forms.Label
    Friend WithEvents txtRiyoNm As System.Windows.Forms.TextBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents rdoStudio As System.Windows.Forms.RadioButton
    Friend WithEvents rdoTheatre As System.Windows.Forms.RadioButton
    Friend WithEvents chkMinyukinOnly As System.Windows.Forms.CheckBox
    Friend WithEvents btnAitesaakiSearch As System.Windows.Forms.Button
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents btnSearch As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtAitesakiNm As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label66 As System.Windows.Forms.Label
    Friend WithEvents btnBack As System.Windows.Forms.Button
    Friend WithEvents vwSeikyuTheatre As FarPoint.Win.Spread.FpSpread
    Friend WithEvents vwSeikyuTheatre_Sheet1 As FarPoint.Win.Spread.SheetView
    Friend WithEvents vwSeikyuStudio As FarPoint.Win.Spread.FpSpread
    Friend WithEvents vwSeikyuStudio_Sheet1 As FarPoint.Win.Spread.SheetView
    Friend WithEvents dtpNyukinDt_To As Common.DateTimePickerEx
    Friend WithEvents dtpNyukinDt_From As Common.DateTimePickerEx
    Friend WithEvents dtpNyukinYoteiDt_To As Common.DateTimePickerEx
    Friend WithEvents dtpNyukinYoteiDt_From As Common.DateTimePickerEx
    Friend WithEvents dtpSeikyuDt_From As Common.DateTimePickerEx
    Friend WithEvents dtpSeikyuDt_To As Common.DateTimePickerEx
    Friend WithEvents Label6 As System.Windows.Forms.Label

End Class
