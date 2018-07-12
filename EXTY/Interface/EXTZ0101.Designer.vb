<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class EXTZ0101
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
        Dim CheckBoxCellType1 As FarPoint.Win.Spread.CellType.CheckBoxCellType = New FarPoint.Win.Spread.CellType.CheckBoxCellType()
        Dim ButtonCellType1 As FarPoint.Win.Spread.CellType.ButtonCellType = New FarPoint.Win.Spread.CellType.ButtonCellType()
        Dim CheckBoxCellType2 As FarPoint.Win.Spread.CellType.CheckBoxCellType = New FarPoint.Win.Spread.CellType.CheckBoxCellType()
        Dim ButtonCellType2 As FarPoint.Win.Spread.CellType.ButtonCellType = New FarPoint.Win.Spread.CellType.ButtonCellType()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(EXTZ0101))
        Me.btnBack = New System.Windows.Forms.Button()
        Me.Label66 = New System.Windows.Forms.Label()
        Me.Label65 = New System.Windows.Forms.Label()
        Me.Label64 = New System.Windows.Forms.Label()
        Me.txtRiyoNm_Kana = New System.Windows.Forms.TextBox()
        Me.Label63 = New System.Windows.Forms.Label()
        Me.txtSaijiShutsuenNm = New System.Windows.Forms.TextBox()
        Me.Label62 = New System.Windows.Forms.Label()
        Me.txtYoyakuNo = New System.Windows.Forms.TextBox()
        Me.Label54 = New System.Windows.Forms.Label()
        Me.Label53 = New System.Windows.Forms.Label()
        Me.txtRiyoNm = New System.Windows.Forms.TextBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.rdoStudio = New System.Windows.Forms.RadioButton()
        Me.rdoTheatre = New System.Windows.Forms.RadioButton()
        Me.chkMikanryoOnly = New System.Windows.Forms.CheckBox()
        Me.btnRiyoshaSearch = New System.Windows.Forms.Button()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.dtpRiyoDt_From = New Common.DateTimePickerEx()
        Me.dtpRiyoDt_To = New Common.DateTimePickerEx()
        Me.btnSearch = New System.Windows.Forms.Button()
        Me.vwYoyakuStudio = New FarPoint.Win.Spread.FpSpread()
        Me.vwYoyakuStudio_Sheet1 = New FarPoint.Win.Spread.SheetView()
        Me.btnDecision = New System.Windows.Forms.Button()
        Me.vwYoyakuTheatre_Sheet1 = New FarPoint.Win.Spread.SheetView()
        Me.vwYoyakuTheatre = New FarPoint.Win.Spread.FpSpread()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Panel1.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        CType(Me.vwYoyakuStudio, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.vwYoyakuStudio_Sheet1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.vwYoyakuTheatre_Sheet1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.vwYoyakuTheatre, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnBack
        '
        Me.btnBack.Location = New System.Drawing.Point(475, 911)
        Me.btnBack.Name = "btnBack"
        Me.btnBack.Size = New System.Drawing.Size(175, 36)
        Me.btnBack.TabIndex = 3
        Me.btnBack.Text = "戻る"
        Me.btnBack.UseVisualStyleBackColor = True
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
        'Label65
        '
        Me.Label65.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label65.Location = New System.Drawing.Point(21, 66)
        Me.Label65.Name = "Label65"
        Me.Label65.Size = New System.Drawing.Size(102, 42)
        Me.Label65.TabIndex = 2
        Me.Label65.Text = "利用日"
        '
        'Label64
        '
        Me.Label64.AutoSize = True
        Me.Label64.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label64.Location = New System.Drawing.Point(21, 99)
        Me.Label64.Name = "Label64"
        Me.Label64.Size = New System.Drawing.Size(82, 13)
        Me.Label64.TabIndex = 8
        Me.Label64.Text = "利用者名(ｶﾅ)" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'txtRiyoNm_Kana
        '
        Me.txtRiyoNm_Kana.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txtRiyoNm_Kana.Location = New System.Drawing.Point(224, 96)
        Me.txtRiyoNm_Kana.Name = "txtRiyoNm_Kana"
        Me.txtRiyoNm_Kana.Size = New System.Drawing.Size(306, 20)
        Me.txtRiyoNm_Kana.TabIndex = 9
        '
        'Label63
        '
        Me.Label63.AutoSize = True
        Me.Label63.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label63.Location = New System.Drawing.Point(21, 137)
        Me.Label63.Name = "Label63"
        Me.Label63.Size = New System.Drawing.Size(129, 13)
        Me.Label63.TabIndex = 11
        Me.Label63.Text = "催事名／アーティスト名"
        '
        'txtSaijiShutsuenNm
        '
        Me.txtSaijiShutsuenNm.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txtSaijiShutsuenNm.Location = New System.Drawing.Point(224, 134)
        Me.txtSaijiShutsuenNm.Name = "txtSaijiShutsuenNm"
        Me.txtSaijiShutsuenNm.Size = New System.Drawing.Size(306, 20)
        Me.txtSaijiShutsuenNm.TabIndex = 12
        '
        'Label62
        '
        Me.Label62.AutoSize = True
        Me.Label62.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label62.Location = New System.Drawing.Point(21, 171)
        Me.Label62.Name = "Label62"
        Me.Label62.Size = New System.Drawing.Size(59, 13)
        Me.Label62.TabIndex = 13
        Me.Label62.Text = "予約番号"
        '
        'txtYoyakuNo
        '
        Me.txtYoyakuNo.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txtYoyakuNo.Location = New System.Drawing.Point(224, 168)
        Me.txtYoyakuNo.Name = "txtYoyakuNo"
        Me.txtYoyakuNo.Size = New System.Drawing.Size(112, 20)
        Me.txtYoyakuNo.TabIndex = 14
        '
        'Label54
        '
        Me.Label54.AutoSize = True
        Me.Label54.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label54.Location = New System.Drawing.Point(652, 65)
        Me.Label54.Name = "Label54"
        Me.Label54.Size = New System.Drawing.Size(59, 13)
        Me.Label54.TabIndex = 6
        Me.Label54.Text = "利用者名"
        '
        'Label53
        '
        Me.Label53.AutoSize = True
        Me.Label53.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label53.Location = New System.Drawing.Point(320, 66)
        Me.Label53.Name = "Label53"
        Me.Label53.Size = New System.Drawing.Size(20, 13)
        Me.Label53.TabIndex = 4
        Me.Label53.Text = "～"
        '
        'txtRiyoNm
        '
        Me.txtRiyoNm.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txtRiyoNm.Location = New System.Drawing.Point(751, 63)
        Me.txtRiyoNm.Name = "txtRiyoNm"
        Me.txtRiyoNm.Size = New System.Drawing.Size(364, 20)
        Me.txtRiyoNm.TabIndex = 7
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
        'chkMikanryoOnly
        '
        Me.chkMikanryoOnly.AutoSize = True
        Me.chkMikanryoOnly.Checked = True
        Me.chkMikanryoOnly.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkMikanryoOnly.Location = New System.Drawing.Point(24, 204)
        Me.chkMikanryoOnly.Name = "chkMikanryoOnly"
        Me.chkMikanryoOnly.Size = New System.Drawing.Size(161, 17)
        Me.chkMikanryoOnly.TabIndex = 15
        Me.chkMikanryoOnly.Text = "未完了の予約のみを表示"
        Me.chkMikanryoOnly.UseVisualStyleBackColor = True
        '
        'btnRiyoshaSearch
        '
        Me.btnRiyoshaSearch.BackgroundImage = Global.EXTY.My.Resources.Resources.マスタ検索
        Me.btnRiyoshaSearch.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnRiyoshaSearch.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.btnRiyoshaSearch.Location = New System.Drawing.Point(536, 92)
        Me.btnRiyoshaSearch.Name = "btnRiyoshaSearch"
        Me.btnRiyoshaSearch.Size = New System.Drawing.Size(36, 27)
        Me.btnRiyoshaSearch.TabIndex = 10
        Me.btnRiyoshaSearch.UseVisualStyleBackColor = True
        '
        'GroupBox5
        '
        Me.GroupBox5.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox5.Controls.Add(Me.dtpRiyoDt_From)
        Me.GroupBox5.Controls.Add(Me.dtpRiyoDt_To)
        Me.GroupBox5.Controls.Add(Me.btnSearch)
        Me.GroupBox5.Controls.Add(Me.btnRiyoshaSearch)
        Me.GroupBox5.Controls.Add(Me.chkMikanryoOnly)
        Me.GroupBox5.Controls.Add(Me.Panel1)
        Me.GroupBox5.Controls.Add(Me.txtRiyoNm)
        Me.GroupBox5.Controls.Add(Me.Label53)
        Me.GroupBox5.Controls.Add(Me.Label54)
        Me.GroupBox5.Controls.Add(Me.txtYoyakuNo)
        Me.GroupBox5.Controls.Add(Me.Label62)
        Me.GroupBox5.Controls.Add(Me.txtSaijiShutsuenNm)
        Me.GroupBox5.Controls.Add(Me.Label63)
        Me.GroupBox5.Controls.Add(Me.txtRiyoNm_Kana)
        Me.GroupBox5.Controls.Add(Me.Label64)
        Me.GroupBox5.Controls.Add(Me.Label65)
        Me.GroupBox5.Controls.Add(Me.Label66)
        Me.GroupBox5.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.GroupBox5.Location = New System.Drawing.Point(12, 42)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(1539, 234)
        Me.GroupBox5.TabIndex = 0
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "検索条件"
        '
        'dtpRiyoDt_From
        '
        Me.dtpRiyoDt_From.CausesValidation = False
        Me.dtpRiyoDt_From.Location = New System.Drawing.Point(170, 63)
        Me.dtpRiyoDt_From.Name = "dtpRiyoDt_From"
        Me.dtpRiyoDt_From.Size = New System.Drawing.Size(147, 30)
        Me.dtpRiyoDt_From.TabIndex = 3
        '
        'dtpRiyoDt_To
        '
        Me.dtpRiyoDt_To.Location = New System.Drawing.Point(352, 62)
        Me.dtpRiyoDt_To.Name = "dtpRiyoDt_To"
        Me.dtpRiyoDt_To.Size = New System.Drawing.Size(149, 28)
        Me.dtpRiyoDt_To.TabIndex = 5
        '
        'btnSearch
        '
        Me.btnSearch.Location = New System.Drawing.Point(1133, 152)
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.Size = New System.Drawing.Size(96, 23)
        Me.btnSearch.TabIndex = 16
        Me.btnSearch.Text = "検索"
        Me.btnSearch.UseVisualStyleBackColor = True
        '
        'vwYoyakuStudio
        '
        Me.vwYoyakuStudio.AccessibleDescription = "vwYoyakuStudio, Sheet1, Row 0, Column 0, "
        Me.vwYoyakuStudio.Location = New System.Drawing.Point(12, 293)
        Me.vwYoyakuStudio.Name = "vwYoyakuStudio"
        Me.vwYoyakuStudio.Sheets.AddRange(New FarPoint.Win.Spread.SheetView() {Me.vwYoyakuStudio_Sheet1})
        Me.vwYoyakuStudio.Size = New System.Drawing.Size(1560, 582)
        Me.vwYoyakuStudio.TabIndex = 1
        '
        'vwYoyakuStudio_Sheet1
        '
        Me.vwYoyakuStudio_Sheet1.Reset()
        Me.vwYoyakuStudio_Sheet1.SheetName = "Sheet1"
        'Formulas and custom names must be loaded with R1C1 reference style
        Me.vwYoyakuStudio_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.R1C1
        Me.vwYoyakuStudio_Sheet1.ColumnCount = 17
        Me.vwYoyakuStudio_Sheet1.RowCount = 1
        Me.vwYoyakuStudio_Sheet1.ColumnHeader.Cells.Get(0, 0).Value = "選択"
        Me.vwYoyakuStudio_Sheet1.ColumnHeader.Cells.Get(0, 1).Value = "ｽﾃｰﾀｽ"
        Me.vwYoyakuStudio_Sheet1.ColumnHeader.Cells.Get(0, 2).Value = "予約番号"
        Me.vwYoyakuStudio_Sheet1.ColumnHeader.Cells.Get(0, 3).Value = "スタジオ"
        Me.vwYoyakuStudio_Sheet1.ColumnHeader.Cells.Get(0, 4).Value = "利用開始日"
        Me.vwYoyakuStudio_Sheet1.ColumnHeader.Cells.Get(0, 5).Value = "利用終了日"
        Me.vwYoyakuStudio_Sheet1.ColumnHeader.Cells.Get(0, 6).Value = "開始時間"
        Me.vwYoyakuStudio_Sheet1.ColumnHeader.Cells.Get(0, 7).Value = "終了時間"
        Me.vwYoyakuStudio_Sheet1.ColumnHeader.Cells.Get(0, 8).Value = "アーティスト名"
        Me.vwYoyakuStudio_Sheet1.ColumnHeader.Cells.Get(0, 9).Value = "利用者名"
        Me.vwYoyakuStudio_Sheet1.ColumnHeader.Cells.Get(0, 10).Value = "責任者名"
        Me.vwYoyakuStudio_Sheet1.ColumnHeader.Cells.Get(0, 11).Value = "承認人数"
        Me.vwYoyakuStudio_Sheet1.ColumnHeader.Cells.Get(0, 12).Value = "利用料"
        Me.vwYoyakuStudio_Sheet1.ColumnHeader.Cells.Get(0, 13).Value = "設備入力ｽﾃｰﾀｽ"
        Me.vwYoyakuStudio_Sheet1.ColumnHeader.Cells.Get(0, 14).Value = "付帯設備"
        Me.vwYoyakuStudio_Sheet1.ColumnHeader.Cells.Get(0, 15).Value = "　"
        Me.vwYoyakuStudio_Sheet1.ColumnHeader.Cells.Get(0, 16).Value = "STS"
        Me.vwYoyakuStudio_Sheet1.ColumnHeader.Rows.Get(0).Height = 32.0!
        Me.vwYoyakuStudio_Sheet1.Columns.Get(0).CellType = CheckBoxCellType1
        Me.vwYoyakuStudio_Sheet1.Columns.Get(0).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwYoyakuStudio_Sheet1.Columns.Get(0).Label = "選択"
        Me.vwYoyakuStudio_Sheet1.Columns.Get(0).Width = 35.0!
        Me.vwYoyakuStudio_Sheet1.Columns.Get(1).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwYoyakuStudio_Sheet1.Columns.Get(1).Label = "ｽﾃｰﾀｽ"
        Me.vwYoyakuStudio_Sheet1.Columns.Get(1).Locked = True
        Me.vwYoyakuStudio_Sheet1.Columns.Get(1).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwYoyakuStudio_Sheet1.Columns.Get(1).Width = 54.0!
        Me.vwYoyakuStudio_Sheet1.Columns.Get(2).AllowAutoSort = True
        Me.vwYoyakuStudio_Sheet1.Columns.Get(2).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwYoyakuStudio_Sheet1.Columns.Get(2).Label = "予約番号"
        Me.vwYoyakuStudio_Sheet1.Columns.Get(2).Locked = True
        Me.vwYoyakuStudio_Sheet1.Columns.Get(2).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwYoyakuStudio_Sheet1.Columns.Get(2).Width = 88.0!
        Me.vwYoyakuStudio_Sheet1.Columns.Get(3).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwYoyakuStudio_Sheet1.Columns.Get(3).Label = "スタジオ"
        Me.vwYoyakuStudio_Sheet1.Columns.Get(3).Locked = True
        Me.vwYoyakuStudio_Sheet1.Columns.Get(3).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwYoyakuStudio_Sheet1.Columns.Get(3).Width = 77.0!
        Me.vwYoyakuStudio_Sheet1.Columns.Get(4).AllowAutoSort = True
        Me.vwYoyakuStudio_Sheet1.Columns.Get(4).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwYoyakuStudio_Sheet1.Columns.Get(4).Label = "利用開始日"
        Me.vwYoyakuStudio_Sheet1.Columns.Get(4).Locked = True
        Me.vwYoyakuStudio_Sheet1.Columns.Get(4).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwYoyakuStudio_Sheet1.Columns.Get(4).Width = 103.0!
        Me.vwYoyakuStudio_Sheet1.Columns.Get(5).AllowAutoSort = True
        Me.vwYoyakuStudio_Sheet1.Columns.Get(5).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwYoyakuStudio_Sheet1.Columns.Get(5).Label = "利用終了日"
        Me.vwYoyakuStudio_Sheet1.Columns.Get(5).Locked = True
        Me.vwYoyakuStudio_Sheet1.Columns.Get(5).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwYoyakuStudio_Sheet1.Columns.Get(5).Width = 103.0!
        Me.vwYoyakuStudio_Sheet1.Columns.Get(6).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwYoyakuStudio_Sheet1.Columns.Get(6).Label = "開始時間"
        Me.vwYoyakuStudio_Sheet1.Columns.Get(6).Locked = True
        Me.vwYoyakuStudio_Sheet1.Columns.Get(6).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwYoyakuStudio_Sheet1.Columns.Get(6).Width = 65.0!
        Me.vwYoyakuStudio_Sheet1.Columns.Get(7).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwYoyakuStudio_Sheet1.Columns.Get(7).Label = "終了時間"
        Me.vwYoyakuStudio_Sheet1.Columns.Get(7).Locked = True
        Me.vwYoyakuStudio_Sheet1.Columns.Get(7).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwYoyakuStudio_Sheet1.Columns.Get(7).Width = 65.0!
        Me.vwYoyakuStudio_Sheet1.Columns.Get(8).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.General
        Me.vwYoyakuStudio_Sheet1.Columns.Get(8).Label = "アーティスト名"
        Me.vwYoyakuStudio_Sheet1.Columns.Get(8).Locked = True
        Me.vwYoyakuStudio_Sheet1.Columns.Get(8).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwYoyakuStudio_Sheet1.Columns.Get(8).Width = 240.0!
        Me.vwYoyakuStudio_Sheet1.Columns.Get(9).AllowAutoSort = True
        Me.vwYoyakuStudio_Sheet1.Columns.Get(9).Label = "利用者名"
        Me.vwYoyakuStudio_Sheet1.Columns.Get(9).Locked = True
        Me.vwYoyakuStudio_Sheet1.Columns.Get(9).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwYoyakuStudio_Sheet1.Columns.Get(9).Width = 157.0!
        Me.vwYoyakuStudio_Sheet1.Columns.Get(10).AllowAutoSort = True
        Me.vwYoyakuStudio_Sheet1.Columns.Get(10).Label = "責任者名"
        Me.vwYoyakuStudio_Sheet1.Columns.Get(10).Locked = True
        Me.vwYoyakuStudio_Sheet1.Columns.Get(10).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwYoyakuStudio_Sheet1.Columns.Get(10).Width = 157.0!
        Me.vwYoyakuStudio_Sheet1.Columns.Get(11).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Right
        Me.vwYoyakuStudio_Sheet1.Columns.Get(11).Label = "承認人数"
        Me.vwYoyakuStudio_Sheet1.Columns.Get(11).Locked = True
        Me.vwYoyakuStudio_Sheet1.Columns.Get(11).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwYoyakuStudio_Sheet1.Columns.Get(11).Width = 68.0!
        Me.vwYoyakuStudio_Sheet1.Columns.Get(12).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwYoyakuStudio_Sheet1.Columns.Get(12).Label = "利用料"
        Me.vwYoyakuStudio_Sheet1.Columns.Get(12).Locked = True
        Me.vwYoyakuStudio_Sheet1.Columns.Get(12).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwYoyakuStudio_Sheet1.Columns.Get(13).Label = "設備入力ｽﾃｰﾀｽ"
        Me.vwYoyakuStudio_Sheet1.Columns.Get(13).Locked = True
        Me.vwYoyakuStudio_Sheet1.Columns.Get(13).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwYoyakuStudio_Sheet1.Columns.Get(14).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwYoyakuStudio_Sheet1.Columns.Get(14).Label = "付帯設備"
        Me.vwYoyakuStudio_Sheet1.Columns.Get(14).Locked = True
        Me.vwYoyakuStudio_Sheet1.Columns.Get(14).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        ButtonCellType1.ButtonColor2 = System.Drawing.SystemColors.ButtonFace
        ButtonCellType1.Text = "確認・編集"
        Me.vwYoyakuStudio_Sheet1.Columns.Get(15).CellType = ButtonCellType1
        Me.vwYoyakuStudio_Sheet1.Columns.Get(15).Label = "　"
        Me.vwYoyakuStudio_Sheet1.Columns.Get(15).Width = 101.0!
        Me.vwYoyakuStudio_Sheet1.Columns.Get(16).Label = "STS"
        Me.vwYoyakuStudio_Sheet1.Columns.Get(16).Visible = False
        Me.vwYoyakuStudio_Sheet1.RowHeader.Columns.Default.Resizable = False
        Me.vwYoyakuStudio_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.A1
        '
        'btnDecision
        '
        Me.btnDecision.Location = New System.Drawing.Point(1029, 911)
        Me.btnDecision.Name = "btnDecision"
        Me.btnDecision.Size = New System.Drawing.Size(175, 36)
        Me.btnDecision.TabIndex = 4
        Me.btnDecision.Text = "選択確定"
        Me.btnDecision.UseVisualStyleBackColor = True
        '
        'vwYoyakuTheatre_Sheet1
        '
        Me.vwYoyakuTheatre_Sheet1.Reset()
        Me.vwYoyakuTheatre_Sheet1.SheetName = "Sheet1"
        'Formulas and custom names must be loaded with R1C1 reference style
        Me.vwYoyakuTheatre_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.R1C1
        Me.vwYoyakuTheatre_Sheet1.ColumnCount = 17
        Me.vwYoyakuTheatre_Sheet1.RowCount = 1
        Me.vwYoyakuTheatre_Sheet1.ColumnHeader.Cells.Get(0, 0).Value = "選択"
        Me.vwYoyakuTheatre_Sheet1.ColumnHeader.Cells.Get(0, 1).Value = "ｽﾃｰﾀｽ"
        Me.vwYoyakuTheatre_Sheet1.ColumnHeader.Cells.Get(0, 2).Value = "予約番号"
        Me.vwYoyakuTheatre_Sheet1.ColumnHeader.Cells.Get(0, 3).Value = "利用開始日"
        Me.vwYoyakuTheatre_Sheet1.ColumnHeader.Cells.Get(0, 4).Value = "利用終了日"
        Me.vwYoyakuTheatre_Sheet1.ColumnHeader.Cells.Get(0, 5).Value = "開始時間"
        Me.vwYoyakuTheatre_Sheet1.ColumnHeader.Cells.Get(0, 6).Value = "終了時間"
        Me.vwYoyakuTheatre_Sheet1.ColumnHeader.Cells.Get(0, 7).Value = "催事名"
        Me.vwYoyakuTheatre_Sheet1.ColumnHeader.Cells.Get(0, 8).Value = "利用形状"
        Me.vwYoyakuTheatre_Sheet1.ColumnHeader.Cells.Get(0, 9).Value = "利用者名"
        Me.vwYoyakuTheatre_Sheet1.ColumnHeader.Cells.Get(0, 10).Value = "責任者名"
        Me.vwYoyakuTheatre_Sheet1.ColumnHeader.Cells.Get(0, 11).Value = "承認人数"
        Me.vwYoyakuTheatre_Sheet1.ColumnHeader.Cells.Get(0, 12).Value = "利用料"
        Me.vwYoyakuTheatre_Sheet1.ColumnHeader.Cells.Get(0, 13).Value = "設備入力ｽﾃｰﾀｽ"
        Me.vwYoyakuTheatre_Sheet1.ColumnHeader.Cells.Get(0, 14).Value = "付帯設備"
        Me.vwYoyakuTheatre_Sheet1.ColumnHeader.Cells.Get(0, 15).Value = "　"
        Me.vwYoyakuTheatre_Sheet1.ColumnHeader.Cells.Get(0, 16).Value = "STS"
        Me.vwYoyakuTheatre_Sheet1.ColumnHeader.Rows.Get(0).Height = 32.0!
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(0).CellType = CheckBoxCellType2
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(0).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(0).Label = "選択"
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(0).Width = 35.0!
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(1).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(1).Label = "ｽﾃｰﾀｽ"
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(1).Locked = True
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(1).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(1).Width = 54.0!
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(2).AllowAutoSort = True
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(2).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(2).Label = "予約番号"
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(2).Locked = True
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(2).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(2).Width = 88.0!
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(3).AllowAutoSort = True
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(3).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(3).Label = "利用開始日"
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(3).Locked = True
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(3).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(3).Width = 103.0!
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(4).AllowAutoSort = True
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(4).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(4).Label = "利用終了日"
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(4).Locked = True
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(4).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(4).Width = 103.0!
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(5).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(5).Label = "開始時間"
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(5).Locked = True
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(5).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(5).Width = 65.0!
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(6).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(6).Label = "終了時間"
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(6).Locked = True
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(6).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(6).Width = 65.0!
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(7).Label = "催事名"
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(7).Locked = True
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(7).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(7).Width = 250.0!
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(8).Label = "利用形状"
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(8).Locked = True
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(8).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(8).Width = 84.0!
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(9).AllowAutoSort = True
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(9).Label = "利用者名"
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(9).Locked = True
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(9).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(9).Width = 157.0!
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(10).AllowAutoSort = True
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(10).Label = "責任者名"
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(10).Locked = True
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(10).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(10).Width = 157.0!
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(11).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Right
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(11).Label = "承認人数"
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(11).Locked = True
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(11).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(11).Width = 68.0!
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(12).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(12).Label = "利用料"
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(12).Locked = True
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(12).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(13).Label = "設備入力ｽﾃｰﾀｽ"
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(13).Locked = True
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(13).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(14).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(14).Label = "付帯設備"
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(14).Locked = True
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(14).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        ButtonCellType2.ButtonColor2 = System.Drawing.SystemColors.ButtonFace
        ButtonCellType2.Text = "確認・編集"
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(15).CellType = ButtonCellType2
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(15).Label = "　"
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(15).Width = 101.0!
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(16).Label = "STS"
        Me.vwYoyakuTheatre_Sheet1.Columns.Get(16).Visible = False
        Me.vwYoyakuTheatre_Sheet1.RowHeader.Columns.Default.Resizable = False
        Me.vwYoyakuTheatre_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.A1
        '
        'vwYoyakuTheatre
        '
        Me.vwYoyakuTheatre.AccessibleDescription = "vwYoyakuTheatre, Sheet1, Row 0, Column 0, "
        Me.vwYoyakuTheatre.Location = New System.Drawing.Point(12, 293)
        Me.vwYoyakuTheatre.Name = "vwYoyakuTheatre"
        Me.vwYoyakuTheatre.Sheets.AddRange(New FarPoint.Win.Spread.SheetView() {Me.vwYoyakuTheatre_Sheet1})
        Me.vwYoyakuTheatre.Size = New System.Drawing.Size(1569, 582)
        Me.vwYoyakuTheatre.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label1.Location = New System.Drawing.Point(21, 884)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(494, 13)
        Me.Label1.TabIndex = 14
        Me.Label1.Text = "※表のデータは、コピーしたい部分を選択し、キーボードの「Ctrl」＋「C」を押すとコピーできます。"
        '
        'EXTZ0101
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.BackColor = System.Drawing.Color.White
        Me.BackgroundImage = Global.EXTY.My.Resources.Resources.背景
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(1584, 962)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnDecision)
        Me.Controls.Add(Me.GroupBox5)
        Me.Controls.Add(Me.btnBack)
        Me.Controls.Add(Me.vwYoyakuTheatre)
        Me.Controls.Add(Me.vwYoyakuStudio)
        Me.DoubleBuffered = True
        Me.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "EXTZ0101"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ちゃり～ん。　予約一覧(共通)"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox5.PerformLayout()
        CType(Me.vwYoyakuStudio, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.vwYoyakuStudio_Sheet1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.vwYoyakuTheatre_Sheet1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.vwYoyakuTheatre, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnBack As System.Windows.Forms.Button
    Friend WithEvents Label66 As System.Windows.Forms.Label
    Friend WithEvents Label65 As System.Windows.Forms.Label
    Friend WithEvents Label64 As System.Windows.Forms.Label
    Friend WithEvents txtRiyoNm_Kana As System.Windows.Forms.TextBox
    Friend WithEvents Label63 As System.Windows.Forms.Label
    Friend WithEvents txtSaijiShutsuenNm As System.Windows.Forms.TextBox
    Friend WithEvents Label62 As System.Windows.Forms.Label
    Friend WithEvents txtYoyakuNo As System.Windows.Forms.TextBox
    Friend WithEvents Label54 As System.Windows.Forms.Label
    Friend WithEvents Label53 As System.Windows.Forms.Label
    Friend WithEvents txtRiyoNm As System.Windows.Forms.TextBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents rdoStudio As System.Windows.Forms.RadioButton
    Friend WithEvents rdoTheatre As System.Windows.Forms.RadioButton
    Friend WithEvents chkMikanryoOnly As System.Windows.Forms.CheckBox
    Friend WithEvents btnRiyoshaSearch As System.Windows.Forms.Button
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents btnSearch As System.Windows.Forms.Button
    Friend WithEvents vwYoyakuStudio As FarPoint.Win.Spread.FpSpread
    Friend WithEvents vwYoyakuStudio_Sheet1 As FarPoint.Win.Spread.SheetView
    Friend WithEvents btnDecision As System.Windows.Forms.Button
    Friend WithEvents dtpRiyoDt_From As Common.DateTimePickerEx
    Friend WithEvents dtpRiyoDt_To As Common.DateTimePickerEx
    Friend WithEvents vwYoyakuTheatre_Sheet1 As FarPoint.Win.Spread.SheetView
    Friend WithEvents vwYoyakuTheatre As FarPoint.Win.Spread.FpSpread
    Friend WithEvents Label1 As System.Windows.Forms.Label

End Class
