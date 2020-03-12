<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class EXTM0102
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
        Dim TextCellType1 As FarPoint.Win.Spread.CellType.TextCellType = New FarPoint.Win.Spread.CellType.TextCellType()
        Dim TextCellType2 As FarPoint.Win.Spread.CellType.TextCellType = New FarPoint.Win.Spread.CellType.TextCellType()
        Dim TextCellType3 As FarPoint.Win.Spread.CellType.TextCellType = New FarPoint.Win.Spread.CellType.TextCellType()
        Dim CheckBoxCellType1 As FarPoint.Win.Spread.CellType.CheckBoxCellType = New FarPoint.Win.Spread.CellType.CheckBoxCellType()
        Dim ComboBoxCellType1 As FarPoint.Win.Spread.CellType.ComboBoxCellType = New FarPoint.Win.Spread.CellType.ComboBoxCellType()
        Dim NumberCellType1 As FarPoint.Win.Spread.CellType.NumberCellType = New FarPoint.Win.Spread.CellType.NumberCellType()
        Dim ComboBoxCellType2 As FarPoint.Win.Spread.CellType.ComboBoxCellType = New FarPoint.Win.Spread.CellType.ComboBoxCellType()
        Dim ComboBoxCellType3 As FarPoint.Win.Spread.CellType.ComboBoxCellType = New FarPoint.Win.Spread.CellType.ComboBoxCellType()
        Dim ComboBoxCellType4 As FarPoint.Win.Spread.CellType.ComboBoxCellType = New FarPoint.Win.Spread.CellType.ComboBoxCellType()
        Dim ComboBoxCellType5 As FarPoint.Win.Spread.CellType.ComboBoxCellType = New FarPoint.Win.Spread.CellType.ComboBoxCellType()
        Dim TextCellType4 As FarPoint.Win.Spread.CellType.TextCellType = New FarPoint.Win.Spread.CellType.TextCellType()
        Dim CheckBoxCellType2 As FarPoint.Win.Spread.CellType.CheckBoxCellType = New FarPoint.Win.Spread.CellType.CheckBoxCellType()
        Dim CheckBoxCellType3 As FarPoint.Win.Spread.CellType.CheckBoxCellType = New FarPoint.Win.Spread.CellType.CheckBoxCellType()
        Dim TextCellType5 As FarPoint.Win.Spread.CellType.TextCellType = New FarPoint.Win.Spread.CellType.TextCellType()
        Dim TextCellType6 As FarPoint.Win.Spread.CellType.TextCellType = New FarPoint.Win.Spread.CellType.TextCellType()
        Dim NumberCellType2 As FarPoint.Win.Spread.CellType.NumberCellType = New FarPoint.Win.Spread.CellType.NumberCellType()
        Dim CheckBoxCellType4 As FarPoint.Win.Spread.CellType.CheckBoxCellType = New FarPoint.Win.Spread.CellType.CheckBoxCellType()
        Dim TextCellType7 As FarPoint.Win.Spread.CellType.TextCellType = New FarPoint.Win.Spread.CellType.TextCellType()
        Dim CheckBoxCellType5 As FarPoint.Win.Spread.CellType.CheckBoxCellType = New FarPoint.Win.Spread.CellType.CheckBoxCellType()
        Dim TextCellType8 As FarPoint.Win.Spread.CellType.TextCellType = New FarPoint.Win.Spread.CellType.TextCellType()
        Dim TextCellType9 As FarPoint.Win.Spread.CellType.TextCellType = New FarPoint.Win.Spread.CellType.TextCellType()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(EXTM0102))
        Me.btnBack = New System.Windows.Forms.Button()
        Me.btnEntry = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.cmbBoxFnishedFromTo = New System.Windows.Forms.ComboBox()
        Me.rdoBtnFinished = New System.Windows.Forms.RadioButton()
        Me.rdoBtnNew = New System.Windows.Forms.RadioButton()
        Me.Label34 = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.rdoBtnStudio = New System.Windows.Forms.RadioButton()
        Me.rdoBtnTheater = New System.Windows.Forms.RadioButton()
        Me.btnNewEntry = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.vwGroupingSheet = New FarPoint.Win.Spread.FpSpread()
        Me.vwGroupingSheet_Sheet1 = New FarPoint.Win.Spread.SheetView()
        Me.vwFutaiSheet = New FarPoint.Win.Spread.FpSpread()
        Me.vwFutaiSheet_Sheet1 = New FarPoint.Win.Spread.SheetView()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtMonthFrom = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.txtYearFrom = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtMonthTo = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtYearTo = New System.Windows.Forms.TextBox()
        Me.btnPeriod = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.vwGroupingSheet, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.vwGroupingSheet_Sheet1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.vwFutaiSheet, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.vwFutaiSheet_Sheet1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnBack
        '
        Me.btnBack.Location = New System.Drawing.Point(412, 1128)
        Me.btnBack.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnBack.Name = "btnBack"
        Me.btnBack.Size = New System.Drawing.Size(152, 36)
        Me.btnBack.TabIndex = 114
        Me.btnBack.Text = "戻る"
        Me.btnBack.UseVisualStyleBackColor = True
        '
        'btnEntry
        '
        Me.btnEntry.Location = New System.Drawing.Point(1057, 1128)
        Me.btnEntry.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnEntry.Name = "btnEntry"
        Me.btnEntry.Size = New System.Drawing.Size(152, 36)
        Me.btnEntry.TabIndex = 113
        Me.btnEntry.Text = "登録"
        Me.btnEntry.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox1.Controls.Add(Me.cmbBoxFnishedFromTo)
        Me.GroupBox1.Controls.Add(Me.rdoBtnFinished)
        Me.GroupBox1.Controls.Add(Me.rdoBtnNew)
        Me.GroupBox1.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(48, 31)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox1.Size = New System.Drawing.Size(1191, 64)
        Me.GroupBox1.TabIndex = 115
        Me.GroupBox1.TabStop = False
        '
        'cmbBoxFnishedFromTo
        '
        Me.cmbBoxFnishedFromTo.FormattingEnabled = True
        Me.cmbBoxFnishedFromTo.Location = New System.Drawing.Point(493, 20)
        Me.cmbBoxFnishedFromTo.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cmbBoxFnishedFromTo.Name = "cmbBoxFnishedFromTo"
        Me.cmbBoxFnishedFromTo.Size = New System.Drawing.Size(443, 24)
        Me.cmbBoxFnishedFromTo.TabIndex = 33
        '
        'rdoBtnFinished
        '
        Me.rdoBtnFinished.AutoSize = True
        Me.rdoBtnFinished.Location = New System.Drawing.Point(259, 21)
        Me.rdoBtnFinished.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.rdoBtnFinished.Name = "rdoBtnFinished"
        Me.rdoBtnFinished.Size = New System.Drawing.Size(217, 21)
        Me.rdoBtnFinished.TabIndex = 32
        Me.rdoBtnFinished.Text = "設定済みの料金を編集する"
        Me.rdoBtnFinished.UseVisualStyleBackColor = True
        '
        'rdoBtnNew
        '
        Me.rdoBtnNew.AutoSize = True
        Me.rdoBtnNew.Location = New System.Drawing.Point(21, 21)
        Me.rdoBtnNew.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.rdoBtnNew.Name = "rdoBtnNew"
        Me.rdoBtnNew.Size = New System.Drawing.Size(184, 21)
        Me.rdoBtnNew.TabIndex = 31
        Me.rdoBtnNew.Text = "新規に料金を設定する"
        Me.rdoBtnNew.UseVisualStyleBackColor = True
        '
        'Label34
        '
        Me.Label34.AutoSize = True
        Me.Label34.BackColor = System.Drawing.Color.Transparent
        Me.Label34.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label34.Location = New System.Drawing.Point(44, 125)
        Me.Label34.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label34.Name = "Label34"
        Me.Label34.Size = New System.Drawing.Size(70, 17)
        Me.Label34.TabIndex = 116
        Me.Label34.Text = "期間　＊"
        '
        'GroupBox2
        '
        Me.GroupBox2.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox2.Controls.Add(Me.rdoBtnStudio)
        Me.GroupBox2.Controls.Add(Me.rdoBtnTheater)
        Me.GroupBox2.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(48, 178)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox2.Size = New System.Drawing.Size(315, 64)
        Me.GroupBox2.TabIndex = 122
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "施設"
        '
        'rdoBtnStudio
        '
        Me.rdoBtnStudio.AutoSize = True
        Me.rdoBtnStudio.Location = New System.Drawing.Point(143, 21)
        Me.rdoBtnStudio.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.rdoBtnStudio.Name = "rdoBtnStudio"
        Me.rdoBtnStudio.Size = New System.Drawing.Size(81, 21)
        Me.rdoBtnStudio.TabIndex = 32
        Me.rdoBtnStudio.Text = "スタジオ"
        Me.rdoBtnStudio.UseVisualStyleBackColor = True
        '
        'rdoBtnTheater
        '
        Me.rdoBtnTheater.AutoSize = True
        Me.rdoBtnTheater.Location = New System.Drawing.Point(21, 21)
        Me.rdoBtnTheater.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.rdoBtnTheater.Name = "rdoBtnTheater"
        Me.rdoBtnTheater.Size = New System.Drawing.Size(82, 21)
        Me.rdoBtnTheater.TabIndex = 31
        Me.rdoBtnTheater.Text = "シアター"
        Me.rdoBtnTheater.UseVisualStyleBackColor = True
        '
        'btnNewEntry
        '
        Me.btnNewEntry.Location = New System.Drawing.Point(669, 1128)
        Me.btnNewEntry.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnNewEntry.Name = "btnNewEntry"
        Me.btnNewEntry.Size = New System.Drawing.Size(263, 36)
        Me.btnNewEntry.TabIndex = 123
        Me.btnNewEntry.Text = "登録内容を元に新規登録"
        Me.btnNewEntry.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label2.Location = New System.Drawing.Point(45, 669)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(92, 15)
        Me.Label2.TabIndex = 138
        Me.Label2.Text = "付帯設備　＊"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.BackColor = System.Drawing.Color.Transparent
        Me.Label14.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label14.Location = New System.Drawing.Point(43, 264)
        Me.Label14.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(62, 15)
        Me.Label14.TabIndex = 137
        Me.Label14.Text = "分類　＊"
        '
        'vwGroupingSheet
        '
        Me.vwGroupingSheet.AccessibleDescription = "vwGroupingSheet, Sheet1, Row 0, Column 0, "
        Me.vwGroupingSheet.Location = New System.Drawing.Point(45, 295)
        Me.vwGroupingSheet.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.vwGroupingSheet.Name = "vwGroupingSheet"
        Me.vwGroupingSheet.Sheets.AddRange(New FarPoint.Win.Spread.SheetView() {Me.vwGroupingSheet_Sheet1})
        Me.vwGroupingSheet.Size = New System.Drawing.Size(1737, 314)
        Me.vwGroupingSheet.TabIndex = 139
        '
        'vwGroupingSheet_Sheet1
        '
        Me.vwGroupingSheet_Sheet1.Reset()
        Me.vwGroupingSheet_Sheet1.SheetName = "Sheet1"
        'Formulas and custom names must be loaded with R1C1 reference style
        Me.vwGroupingSheet_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.R1C1
        Me.vwGroupingSheet_Sheet1.ColumnCount = 13
        Me.vwGroupingSheet_Sheet1.ColumnHeader.Cells.Get(0, 0).Value = "分類" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "コード"
        Me.vwGroupingSheet_Sheet1.ColumnHeader.Cells.Get(0, 1).Value = "分類名"
        Me.vwGroupingSheet_Sheet1.ColumnHeader.Cells.Get(0, 2).Value = "集計キー"
        Me.vwGroupingSheet_Sheet1.ColumnHeader.Cells.Get(0, 3).Value = "内税"
        Me.vwGroupingSheet_Sheet1.ColumnHeader.Cells.Get(0, 4).Value = "税区分"
        Me.vwGroupingSheet_Sheet1.ColumnHeader.Cells.Get(0, 5).Value = "税率"
        Me.vwGroupingSheet_Sheet1.ColumnHeader.Cells.Get(0, 6).Value = "勘定科目"
        Me.vwGroupingSheet_Sheet1.ColumnHeader.Cells.Get(0, 7).Value = "細目"
        Me.vwGroupingSheet_Sheet1.ColumnHeader.Cells.Get(0, 8).Value = "内訳"
        Me.vwGroupingSheet_Sheet1.ColumnHeader.Cells.Get(0, 9).Value = "詳細"
        Me.vwGroupingSheet_Sheet1.ColumnHeader.Cells.Get(0, 10).Value = "並び順"
        Me.vwGroupingSheet_Sheet1.ColumnHeader.Cells.Get(0, 11).Value = "無効"
        Me.vwGroupingSheet_Sheet1.ColumnHeader.Cells.Get(0, 12).Value = "付帯利用料" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "フラグ"
        Me.vwGroupingSheet_Sheet1.ColumnHeader.Rows.Get(0).Height = 31.0!
        Me.vwGroupingSheet_Sheet1.Columns.Get(0).BackColor = System.Drawing.Color.FromArgb(CType(CType(246, Byte), Integer), CType(CType(234, Byte), Integer), CType(CType(210, Byte), Integer))
        TextCellType1.MaxLength = 2
        Me.vwGroupingSheet_Sheet1.Columns.Get(0).CellType = TextCellType1
        Me.vwGroupingSheet_Sheet1.Columns.Get(0).ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.vwGroupingSheet_Sheet1.Columns.Get(0).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwGroupingSheet_Sheet1.Columns.Get(0).Label = "分類" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "コード"
        Me.vwGroupingSheet_Sheet1.Columns.Get(0).Locked = True
        Me.vwGroupingSheet_Sheet1.Columns.Get(0).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwGroupingSheet_Sheet1.Columns.Get(0).Width = 42.0!
        TextCellType2.MaxLength = 20
        Me.vwGroupingSheet_Sheet1.Columns.Get(1).CellType = TextCellType2
        Me.vwGroupingSheet_Sheet1.Columns.Get(1).Label = "分類名"
        Me.vwGroupingSheet_Sheet1.Columns.Get(1).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwGroupingSheet_Sheet1.Columns.Get(1).Width = 190.0!
        TextCellType3.MaxLength = 20
        Me.vwGroupingSheet_Sheet1.Columns.Get(2).CellType = TextCellType3
        Me.vwGroupingSheet_Sheet1.Columns.Get(2).Label = "集計キー"
        Me.vwGroupingSheet_Sheet1.Columns.Get(2).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwGroupingSheet_Sheet1.Columns.Get(2).Width = 164.0!
        Me.vwGroupingSheet_Sheet1.Columns.Get(3).CellType = CheckBoxCellType1
        Me.vwGroupingSheet_Sheet1.Columns.Get(3).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwGroupingSheet_Sheet1.Columns.Get(3).Label = "内税"
        Me.vwGroupingSheet_Sheet1.Columns.Get(3).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwGroupingSheet_Sheet1.Columns.Get(3).Width = 40.0!
        ComboBoxCellType1.AllowEditorVerticalAlign = True
        ComboBoxCellType1.ButtonAlign = FarPoint.Win.ButtonAlign.Right
        Me.vwGroupingSheet_Sheet1.Columns.Get(4).CellType = ComboBoxCellType1
        Me.vwGroupingSheet_Sheet1.Columns.Get(4).Label = "税区分"
        Me.vwGroupingSheet_Sheet1.Columns.Get(4).Width = 69.0!
        Me.vwGroupingSheet_Sheet1.Columns.Get(5).CellType = NumberCellType1
        Me.vwGroupingSheet_Sheet1.Columns.Get(5).Label = "税率"
        Me.vwGroupingSheet_Sheet1.Columns.Get(5).Locked = True
        ComboBoxCellType2.AllowEditorVerticalAlign = True
        ComboBoxCellType2.ButtonAlign = FarPoint.Win.ButtonAlign.Right
        Me.vwGroupingSheet_Sheet1.Columns.Get(6).CellType = ComboBoxCellType2
        Me.vwGroupingSheet_Sheet1.Columns.Get(6).Label = "勘定科目"
        Me.vwGroupingSheet_Sheet1.Columns.Get(6).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwGroupingSheet_Sheet1.Columns.Get(6).Width = 150.0!
        ComboBoxCellType3.AllowEditorVerticalAlign = True
        ComboBoxCellType3.ButtonAlign = FarPoint.Win.ButtonAlign.Right
        Me.vwGroupingSheet_Sheet1.Columns.Get(7).CellType = ComboBoxCellType3
        Me.vwGroupingSheet_Sheet1.Columns.Get(7).Label = "細目"
        Me.vwGroupingSheet_Sheet1.Columns.Get(7).Locked = True
        Me.vwGroupingSheet_Sheet1.Columns.Get(7).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwGroupingSheet_Sheet1.Columns.Get(7).Width = 150.0!
        ComboBoxCellType4.AllowEditorVerticalAlign = True
        ComboBoxCellType4.ButtonAlign = FarPoint.Win.ButtonAlign.Right
        Me.vwGroupingSheet_Sheet1.Columns.Get(8).CellType = ComboBoxCellType4
        Me.vwGroupingSheet_Sheet1.Columns.Get(8).Label = "内訳"
        Me.vwGroupingSheet_Sheet1.Columns.Get(8).Locked = True
        Me.vwGroupingSheet_Sheet1.Columns.Get(8).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwGroupingSheet_Sheet1.Columns.Get(8).Width = 150.0!
        ComboBoxCellType5.AllowEditorVerticalAlign = True
        ComboBoxCellType5.ButtonAlign = FarPoint.Win.ButtonAlign.Right
        Me.vwGroupingSheet_Sheet1.Columns.Get(9).CellType = ComboBoxCellType5
        Me.vwGroupingSheet_Sheet1.Columns.Get(9).Label = "詳細"
        Me.vwGroupingSheet_Sheet1.Columns.Get(9).Locked = True
        Me.vwGroupingSheet_Sheet1.Columns.Get(9).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwGroupingSheet_Sheet1.Columns.Get(9).Width = 150.0!
        TextCellType4.MaxLength = 4
        Me.vwGroupingSheet_Sheet1.Columns.Get(10).CellType = TextCellType4
        Me.vwGroupingSheet_Sheet1.Columns.Get(10).Label = "並び順"
        Me.vwGroupingSheet_Sheet1.Columns.Get(10).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwGroupingSheet_Sheet1.Columns.Get(10).Width = 50.0!
        Me.vwGroupingSheet_Sheet1.Columns.Get(11).CellType = CheckBoxCellType2
        Me.vwGroupingSheet_Sheet1.Columns.Get(11).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwGroupingSheet_Sheet1.Columns.Get(11).Label = "無効"
        Me.vwGroupingSheet_Sheet1.Columns.Get(11).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwGroupingSheet_Sheet1.Columns.Get(11).Width = 40.0!
        Me.vwGroupingSheet_Sheet1.Columns.Get(12).CellType = CheckBoxCellType3
        Me.vwGroupingSheet_Sheet1.Columns.Get(12).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwGroupingSheet_Sheet1.Columns.Get(12).Label = "付帯利用料" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "フラグ"
        Me.vwGroupingSheet_Sheet1.Columns.Get(12).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwGroupingSheet_Sheet1.Columns.Get(12).Width = 69.0!
        Me.vwGroupingSheet_Sheet1.RestrictRows = True
        Me.vwGroupingSheet_Sheet1.RowHeader.Columns.Default.Resizable = False
        Me.vwGroupingSheet_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.A1
        '
        'vwFutaiSheet
        '
        Me.vwFutaiSheet.AccessibleDescription = "vwFutaiSheet, Sheet1, Row 0, Column 0, "
        Me.vwFutaiSheet.Location = New System.Drawing.Point(45, 699)
        Me.vwFutaiSheet.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.vwFutaiSheet.Name = "vwFutaiSheet"
        Me.vwFutaiSheet.Sheets.AddRange(New FarPoint.Win.Spread.SheetView() {Me.vwFutaiSheet_Sheet1})
        Me.vwFutaiSheet.Size = New System.Drawing.Size(929, 393)
        Me.vwFutaiSheet.TabIndex = 140
        '
        'vwFutaiSheet_Sheet1
        '
        Me.vwFutaiSheet_Sheet1.Reset()
        Me.vwFutaiSheet_Sheet1.SheetName = "Sheet1"
        'Formulas and custom names must be loaded with R1C1 reference style
        Me.vwFutaiSheet_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.R1C1
        Me.vwFutaiSheet_Sheet1.ColumnCount = 10
        Me.vwFutaiSheet_Sheet1.Cells.Get(0, 0).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwFutaiSheet_Sheet1.Cells.Get(0, 1).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwFutaiSheet_Sheet1.Cells.Get(0, 2).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwFutaiSheet_Sheet1.ColumnFooter.DefaultStyle.HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.General
        Me.vwFutaiSheet_Sheet1.ColumnFooter.DefaultStyle.NoteIndicatorColor = System.Drawing.Color.Red
        Me.vwFutaiSheet_Sheet1.ColumnFooter.DefaultStyle.Parent = "ColumnHeaderDefaultEnhanced"
        Me.vwFutaiSheet_Sheet1.ColumnFooter.DefaultStyle.VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.General
        Me.vwFutaiSheet_Sheet1.ColumnFooterSheetCornerStyle.HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.General
        Me.vwFutaiSheet_Sheet1.ColumnFooterSheetCornerStyle.NoteIndicatorColor = System.Drawing.Color.Red
        Me.vwFutaiSheet_Sheet1.ColumnFooterSheetCornerStyle.Parent = "CornerFooterDefaultEnhanced"
        Me.vwFutaiSheet_Sheet1.ColumnFooterSheetCornerStyle.VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.General
        Me.vwFutaiSheet_Sheet1.ColumnHeader.Cells.Get(0, 0).Value = "設備" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "コード"
        Me.vwFutaiSheet_Sheet1.ColumnHeader.Cells.Get(0, 1).Value = "設備名"
        Me.vwFutaiSheet_Sheet1.ColumnHeader.Cells.Get(0, 2).Value = "単価"
        Me.vwFutaiSheet_Sheet1.ColumnHeader.Cells.Get(0, 3).Value = "単位"
        Me.vwFutaiSheet_Sheet1.ColumnHeader.Cells.Get(0, 4).Value = "デフォルト" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "セット"
        Me.vwFutaiSheet_Sheet1.ColumnHeader.Cells.Get(0, 5).Value = "並び順"
        Me.vwFutaiSheet_Sheet1.ColumnHeader.Cells.Get(0, 6).Value = "無効"
        Me.vwFutaiSheet_Sheet1.ColumnHeader.Cells.Get(0, 7).Value = "行番号"
        Me.vwFutaiSheet_Sheet1.ColumnHeader.Cells.Get(0, 8).Value = "分類コード"
        Me.vwFutaiSheet_Sheet1.ColumnHeader.Cells.Get(0, 9).Value = "更新区分"
        Me.vwFutaiSheet_Sheet1.ColumnHeader.DefaultStyle.HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.General
        Me.vwFutaiSheet_Sheet1.ColumnHeader.DefaultStyle.NoteIndicatorColor = System.Drawing.Color.Red
        Me.vwFutaiSheet_Sheet1.ColumnHeader.DefaultStyle.Parent = "ColumnHeaderDefaultEnhanced"
        Me.vwFutaiSheet_Sheet1.ColumnHeader.DefaultStyle.VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.General
        Me.vwFutaiSheet_Sheet1.ColumnHeader.Rows.Get(0).Height = 33.0!
        Me.vwFutaiSheet_Sheet1.Columns.Get(0).BackColor = System.Drawing.Color.FromArgb(CType(CType(246, Byte), Integer), CType(CType(234, Byte), Integer), CType(CType(210, Byte), Integer))
        TextCellType5.MaxLength = 3
        Me.vwFutaiSheet_Sheet1.Columns.Get(0).CellType = TextCellType5
        Me.vwFutaiSheet_Sheet1.Columns.Get(0).Label = "設備" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "コード"
        Me.vwFutaiSheet_Sheet1.Columns.Get(0).Locked = True
        Me.vwFutaiSheet_Sheet1.Columns.Get(0).VisualStyles = FarPoint.Win.VisualStyles.[Auto]
        TextCellType6.MaxLength = 32
        Me.vwFutaiSheet_Sheet1.Columns.Get(1).CellType = TextCellType6
        Me.vwFutaiSheet_Sheet1.Columns.Get(1).Label = "設備名"
        Me.vwFutaiSheet_Sheet1.Columns.Get(1).Width = 287.0!
        NumberCellType2.DecimalPlaces = 0
        NumberCellType2.FixedPoint = False
        NumberCellType2.MaximumValue = 99999999.0R
        NumberCellType2.MinimumValue = 0R
        NumberCellType2.ShowSeparator = True
        Me.vwFutaiSheet_Sheet1.Columns.Get(2).CellType = NumberCellType2
        Me.vwFutaiSheet_Sheet1.Columns.Get(2).Label = "単価"
        Me.vwFutaiSheet_Sheet1.Columns.Get(2).Width = 65.0!
        Me.vwFutaiSheet_Sheet1.Columns.Get(3).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwFutaiSheet_Sheet1.Columns.Get(3).Label = "単位"
        Me.vwFutaiSheet_Sheet1.Columns.Get(3).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwFutaiSheet_Sheet1.Columns.Get(3).Width = 65.0!
        Me.vwFutaiSheet_Sheet1.Columns.Get(4).CellType = CheckBoxCellType4
        Me.vwFutaiSheet_Sheet1.Columns.Get(4).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwFutaiSheet_Sheet1.Columns.Get(4).Label = "デフォルト" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "セット"
        Me.vwFutaiSheet_Sheet1.Columns.Get(4).Width = 61.0!
        TextCellType7.MaxLength = 4
        Me.vwFutaiSheet_Sheet1.Columns.Get(5).CellType = TextCellType7
        Me.vwFutaiSheet_Sheet1.Columns.Get(5).Label = "並び順"
        Me.vwFutaiSheet_Sheet1.Columns.Get(5).Width = 43.0!
        Me.vwFutaiSheet_Sheet1.Columns.Get(6).CellType = CheckBoxCellType5
        Me.vwFutaiSheet_Sheet1.Columns.Get(6).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwFutaiSheet_Sheet1.Columns.Get(6).Label = "無効"
        Me.vwFutaiSheet_Sheet1.Columns.Get(7).Label = "行番号"
        Me.vwFutaiSheet_Sheet1.Columns.Get(7).Visible = False
        Me.vwFutaiSheet_Sheet1.Columns.Get(8).CellType = TextCellType8
        Me.vwFutaiSheet_Sheet1.Columns.Get(8).Label = "分類コード"
        Me.vwFutaiSheet_Sheet1.Columns.Get(8).Visible = False
        Me.vwFutaiSheet_Sheet1.Columns.Get(9).CellType = TextCellType9
        Me.vwFutaiSheet_Sheet1.Columns.Get(9).Label = "更新区分"
        Me.vwFutaiSheet_Sheet1.Columns.Get(9).Visible = False
        Me.vwFutaiSheet_Sheet1.DefaultStyle.HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.General
        Me.vwFutaiSheet_Sheet1.DefaultStyle.NoteIndicatorColor = System.Drawing.Color.Red
        Me.vwFutaiSheet_Sheet1.DefaultStyle.Parent = "DataAreaDefault"
        Me.vwFutaiSheet_Sheet1.DefaultStyle.VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.General
        Me.vwFutaiSheet_Sheet1.FilterBar.DefaultStyle.HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.General
        Me.vwFutaiSheet_Sheet1.FilterBar.DefaultStyle.NoteIndicatorColor = System.Drawing.Color.Red
        Me.vwFutaiSheet_Sheet1.FilterBar.DefaultStyle.Parent = "FilterBarDefaultEnhanced"
        Me.vwFutaiSheet_Sheet1.FilterBar.DefaultStyle.VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.General
        Me.vwFutaiSheet_Sheet1.FilterBarHeaderStyle.HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.General
        Me.vwFutaiSheet_Sheet1.FilterBarHeaderStyle.NoteIndicatorColor = System.Drawing.Color.Red
        Me.vwFutaiSheet_Sheet1.FilterBarHeaderStyle.Parent = "RowHeaderDefaultEnhanced"
        Me.vwFutaiSheet_Sheet1.FilterBarHeaderStyle.VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.General
        Me.vwFutaiSheet_Sheet1.RowHeader.Columns.Default.Resizable = False
        Me.vwFutaiSheet_Sheet1.RowHeader.DefaultStyle.HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.General
        Me.vwFutaiSheet_Sheet1.RowHeader.DefaultStyle.NoteIndicatorColor = System.Drawing.Color.Red
        Me.vwFutaiSheet_Sheet1.RowHeader.DefaultStyle.Parent = "RowHeaderDefaultEnhanced"
        Me.vwFutaiSheet_Sheet1.RowHeader.DefaultStyle.VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.General
        Me.vwFutaiSheet_Sheet1.SheetCornerStyle.HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.General
        Me.vwFutaiSheet_Sheet1.SheetCornerStyle.NoteIndicatorColor = System.Drawing.Color.Red
        Me.vwFutaiSheet_Sheet1.SheetCornerStyle.Parent = "CornerDefaultEnhanced"
        Me.vwFutaiSheet_Sheet1.SheetCornerStyle.VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.General
        Me.vwFutaiSheet_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.A1
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.BackColor = System.Drawing.Color.Transparent
        Me.Label5.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label5.Location = New System.Drawing.Point(299, 124)
        Me.Label5.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(47, 15)
        Me.Label5.TabIndex = 144
        Me.Label5.Text = "月　～"
        '
        'txtMonthFrom
        '
        Me.txtMonthFrom.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txtMonthFrom.Location = New System.Drawing.Point(243, 120)
        Me.txtMonthFrom.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtMonthFrom.MaxLength = 2
        Me.txtMonthFrom.Name = "txtMonthFrom"
        Me.txtMonthFrom.Size = New System.Drawing.Size(47, 22)
        Me.txtMonthFrom.TabIndex = 143
        Me.txtMonthFrom.Text = "5"
        Me.txtMonthFrom.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.BackColor = System.Drawing.Color.Transparent
        Me.Label11.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label11.Location = New System.Drawing.Point(212, 124)
        Me.Label11.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(22, 15)
        Me.Label11.TabIndex = 142
        Me.Label11.Text = "年"
        '
        'txtYearFrom
        '
        Me.txtYearFrom.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txtYearFrom.Location = New System.Drawing.Point(125, 120)
        Me.txtYearFrom.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtYearFrom.MaxLength = 4
        Me.txtYearFrom.Name = "txtYearFrom"
        Me.txtYearFrom.Size = New System.Drawing.Size(77, 22)
        Me.txtYearFrom.TabIndex = 141
        Me.txtYearFrom.Text = "2015"
        Me.txtYearFrom.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label1.Location = New System.Drawing.Point(545, 124)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(22, 15)
        Me.Label1.TabIndex = 148
        Me.Label1.Text = "月"
        '
        'txtMonthTo
        '
        Me.txtMonthTo.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txtMonthTo.Location = New System.Drawing.Point(489, 120)
        Me.txtMonthTo.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtMonthTo.MaxLength = 2
        Me.txtMonthTo.Name = "txtMonthTo"
        Me.txtMonthTo.Size = New System.Drawing.Size(47, 22)
        Me.txtMonthTo.TabIndex = 147
        Me.txtMonthTo.Text = "5"
        Me.txtMonthTo.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label3.Location = New System.Drawing.Point(459, 124)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(22, 15)
        Me.Label3.TabIndex = 146
        Me.Label3.Text = "年"
        '
        'txtYearTo
        '
        Me.txtYearTo.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txtYearTo.Location = New System.Drawing.Point(372, 120)
        Me.txtYearTo.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtYearTo.MaxLength = 4
        Me.txtYearTo.Name = "txtYearTo"
        Me.txtYearTo.Size = New System.Drawing.Size(77, 22)
        Me.txtYearTo.TabIndex = 145
        Me.txtYearTo.Text = "2015"
        Me.txtYearTo.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'btnPeriod
        '
        Me.btnPeriod.Location = New System.Drawing.Point(576, 120)
        Me.btnPeriod.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnPeriod.Name = "btnPeriod"
        Me.btnPeriod.Size = New System.Drawing.Size(120, 24)
        Me.btnPeriod.TabIndex = 149
        Me.btnPeriod.Text = "設定"
        Me.btnPeriod.UseVisualStyleBackColor = True
        '
        'EXTM0102
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.BackColor = System.Drawing.Color.White
        Me.BackgroundImage = Global.EXTM.My.Resources.Resources.背景
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(1835, 1185)
        Me.Controls.Add(Me.btnPeriod)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtMonthTo)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtYearTo)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.txtMonthFrom)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.txtYearFrom)
        Me.Controls.Add(Me.vwFutaiSheet)
        Me.Controls.Add(Me.vwGroupingSheet)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.btnNewEntry)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.Label34)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.btnBack)
        Me.Controls.Add(Me.btnEntry)
        Me.DoubleBuffered = True
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Name = "EXTM0102"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ちゃり～ん。　付帯設備マスタメンテ"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.vwGroupingSheet, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.vwGroupingSheet_Sheet1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.vwFutaiSheet, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.vwFutaiSheet_Sheet1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnBack As System.Windows.Forms.Button
    Friend WithEvents btnEntry As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents cmbBoxFnishedFromTo As System.Windows.Forms.ComboBox
    Friend WithEvents rdoBtnFinished As System.Windows.Forms.RadioButton
    Friend WithEvents rdoBtnNew As System.Windows.Forms.RadioButton
    Friend WithEvents Label34 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents rdoBtnStudio As System.Windows.Forms.RadioButton
    Friend WithEvents rdoBtnTheater As System.Windows.Forms.RadioButton
    Friend WithEvents btnNewEntry As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents vwGroupingSheet As FarPoint.Win.Spread.FpSpread
    Friend WithEvents vwGroupingSheet_Sheet1 As FarPoint.Win.Spread.SheetView
    Friend WithEvents vwFutaiSheet As FarPoint.Win.Spread.FpSpread
    Friend WithEvents vwFutaiSheet_Sheet1 As FarPoint.Win.Spread.SheetView
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtMonthFrom As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents txtYearFrom As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtMonthTo As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtYearTo As System.Windows.Forms.TextBox
    Friend WithEvents btnPeriod As Button
End Class
