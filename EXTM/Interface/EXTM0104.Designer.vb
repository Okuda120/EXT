<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class EXTM0104
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
        Dim ComboBoxCellType1 As FarPoint.Win.Spread.CellType.ComboBoxCellType = New FarPoint.Win.Spread.CellType.ComboBoxCellType()
        Dim TextCellType2 As FarPoint.Win.Spread.CellType.TextCellType = New FarPoint.Win.Spread.CellType.TextCellType()
        Dim TextCellType3 As FarPoint.Win.Spread.CellType.TextCellType = New FarPoint.Win.Spread.CellType.TextCellType()
        Dim CheckBoxCellType1 As FarPoint.Win.Spread.CellType.CheckBoxCellType = New FarPoint.Win.Spread.CellType.CheckBoxCellType()
        Dim TextCellType4 As FarPoint.Win.Spread.CellType.TextCellType = New FarPoint.Win.Spread.CellType.TextCellType()
        Dim TextCellType5 As FarPoint.Win.Spread.CellType.TextCellType = New FarPoint.Win.Spread.CellType.TextCellType()
        Dim TextCellType6 As FarPoint.Win.Spread.CellType.TextCellType = New FarPoint.Win.Spread.CellType.TextCellType()
        Dim NumberCellType1 As FarPoint.Win.Spread.CellType.NumberCellType = New FarPoint.Win.Spread.CellType.NumberCellType()
        Dim NumberCellType2 As FarPoint.Win.Spread.CellType.NumberCellType = New FarPoint.Win.Spread.CellType.NumberCellType()
        Dim TextCellType7 As FarPoint.Win.Spread.CellType.TextCellType = New FarPoint.Win.Spread.CellType.TextCellType()
        Dim CheckBoxCellType2 As FarPoint.Win.Spread.CellType.CheckBoxCellType = New FarPoint.Win.Spread.CellType.CheckBoxCellType()
        Dim TextCellType8 As FarPoint.Win.Spread.CellType.TextCellType = New FarPoint.Win.Spread.CellType.TextCellType()
        Dim ComboBoxCellType2 As FarPoint.Win.Spread.CellType.ComboBoxCellType = New FarPoint.Win.Spread.CellType.ComboBoxCellType()
        Dim TextCellType9 As FarPoint.Win.Spread.CellType.TextCellType = New FarPoint.Win.Spread.CellType.TextCellType()
        Dim TextCellType10 As FarPoint.Win.Spread.CellType.TextCellType = New FarPoint.Win.Spread.CellType.TextCellType()
        Dim CheckBoxCellType3 As FarPoint.Win.Spread.CellType.CheckBoxCellType = New FarPoint.Win.Spread.CellType.CheckBoxCellType()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(EXTM0104))
        Me.btnsinki = New System.Windows.Forms.Button()
        Me.Label34 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.cmbkikan = New System.Windows.Forms.ComboBox()
        Me.rdosumi = New System.Windows.Forms.RadioButton()
        Me.rdoshinki = New System.Windows.Forms.RadioButton()
        Me.btnback = New System.Windows.Forms.Button()
        Me.btninsert = New System.Windows.Forms.Button()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.vwbunrui = New FarPoint.Win.Spread.FpSpread()
        Me.vwbunrui_Sheet1 = New FarPoint.Win.Spread.SheetView()
        Me.vwryokin = New FarPoint.Win.Spread.FpSpread()
        Me.vwryokin_Sheet1 = New FarPoint.Win.Spread.SheetView()
        Me.vwbairitu = New FarPoint.Win.Spread.FpSpread()
        Me.vwbairitu_Sheet1 = New FarPoint.Win.Spread.SheetView()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtkikantomonth = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtkikantoyear = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtkikanfrommonth = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.txtkikanfromyear = New System.Windows.Forms.TextBox()
        Me.GroupBox1.SuspendLayout()
        CType(Me.vwbunrui, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.vwbunrui_Sheet1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.vwryokin, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.vwryokin_Sheet1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.vwbairitu, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.vwbairitu_Sheet1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnsinki
        '
        Me.btnsinki.Location = New System.Drawing.Point(768, 1129)
        Me.btnsinki.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnsinki.Name = "btnsinki"
        Me.btnsinki.Size = New System.Drawing.Size(263, 36)
        Me.btnsinki.TabIndex = 134
        Me.btnsinki.Text = "登録内容を元に新規登録"
        Me.btnsinki.UseVisualStyleBackColor = True
        '
        'Label34
        '
        Me.Label34.AutoSize = True
        Me.Label34.BackColor = System.Drawing.Color.Transparent
        Me.Label34.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label34.Location = New System.Drawing.Point(55, 124)
        Me.Label34.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label34.Name = "Label34"
        Me.Label34.Size = New System.Drawing.Size(70, 17)
        Me.Label34.TabIndex = 127
        Me.Label34.Text = "期間　＊"
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox1.Controls.Add(Me.cmbkikan)
        Me.GroupBox1.Controls.Add(Me.rdosumi)
        Me.GroupBox1.Controls.Add(Me.rdoshinki)
        Me.GroupBox1.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(59, 32)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox1.Size = New System.Drawing.Size(1191, 64)
        Me.GroupBox1.TabIndex = 126
        Me.GroupBox1.TabStop = False
        '
        'cmbkikan
        '
        Me.cmbkikan.FormattingEnabled = True
        Me.cmbkikan.Location = New System.Drawing.Point(493, 20)
        Me.cmbkikan.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cmbkikan.Name = "cmbkikan"
        Me.cmbkikan.Size = New System.Drawing.Size(443, 24)
        Me.cmbkikan.TabIndex = 33
        '
        'rdosumi
        '
        Me.rdosumi.AutoSize = True
        Me.rdosumi.Location = New System.Drawing.Point(259, 21)
        Me.rdosumi.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.rdosumi.Name = "rdosumi"
        Me.rdosumi.Size = New System.Drawing.Size(217, 21)
        Me.rdosumi.TabIndex = 32
        Me.rdosumi.Text = "設定済みの料金を編集する"
        Me.rdosumi.UseVisualStyleBackColor = True
        '
        'rdoshinki
        '
        Me.rdoshinki.AutoSize = True
        Me.rdoshinki.Location = New System.Drawing.Point(21, 21)
        Me.rdoshinki.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.rdoshinki.Name = "rdoshinki"
        Me.rdoshinki.Size = New System.Drawing.Size(184, 21)
        Me.rdoshinki.TabIndex = 31
        Me.rdoshinki.Text = "新規に料金を設定する"
        Me.rdoshinki.UseVisualStyleBackColor = True
        '
        'btnback
        '
        Me.btnback.Location = New System.Drawing.Point(423, 1129)
        Me.btnback.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnback.Name = "btnback"
        Me.btnback.Size = New System.Drawing.Size(152, 36)
        Me.btnback.TabIndex = 125
        Me.btnback.Text = "戻る"
        Me.btnback.UseVisualStyleBackColor = True
        '
        'btninsert
        '
        Me.btninsert.Location = New System.Drawing.Point(1196, 1129)
        Me.btninsert.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btninsert.Name = "btninsert"
        Me.btninsert.Size = New System.Drawing.Size(152, 36)
        Me.btninsert.TabIndex = 124
        Me.btninsert.Text = "登録"
        Me.btninsert.UseVisualStyleBackColor = True
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.BackColor = System.Drawing.Color.Transparent
        Me.Label14.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label14.Location = New System.Drawing.Point(56, 201)
        Me.Label14.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(123, 15)
        Me.Label14.TabIndex = 135
        Me.Label14.Text = "利用料（分類）　＊"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label2.Location = New System.Drawing.Point(959, 201)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(123, 15)
        Me.Label2.TabIndex = 136
        Me.Label2.Text = "利用料（料金）　＊"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label3.Location = New System.Drawing.Point(56, 686)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(62, 15)
        Me.Label3.TabIndex = 137
        Me.Label3.Text = "倍率　＊"
        '
        'vwbunrui
        '
        Me.vwbunrui.AccessibleDescription = "vwbunrui, Sheet1, Row 0, Column 0, "
        Me.vwbunrui.Location = New System.Drawing.Point(59, 234)
        Me.vwbunrui.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.vwbunrui.Name = "vwbunrui"
        Me.vwbunrui.Sheets.AddRange(New FarPoint.Win.Spread.SheetView() {Me.vwbunrui_Sheet1})
        Me.vwbunrui.Size = New System.Drawing.Size(853, 386)
        Me.vwbunrui.TabIndex = 138
        Me.vwbunrui.SetViewportTopRow(0, 0, 3)
        '
        'vwbunrui_Sheet1
        '
        Me.vwbunrui_Sheet1.Reset()
        Me.vwbunrui_Sheet1.SheetName = "Sheet1"
        'Formulas and custom names must be loaded with R1C1 reference style
        Me.vwbunrui_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.R1C1
        Me.vwbunrui_Sheet1.ColumnCount = 5
        Me.vwbunrui_Sheet1.ColumnHeader.Cells.Get(0, 0).Value = "料金分類" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "コード"
        Me.vwbunrui_Sheet1.ColumnHeader.Cells.Get(0, 1).Value = "施設"
        Me.vwbunrui_Sheet1.ColumnHeader.Cells.Get(0, 2).Value = "料金分類名"
        Me.vwbunrui_Sheet1.ColumnHeader.Cells.Get(0, 3).Value = "並び順"
        Me.vwbunrui_Sheet1.ColumnHeader.Cells.Get(0, 4).Value = "無効"
        Me.vwbunrui_Sheet1.ColumnHeader.Rows.Get(0).Height = 32.0!
        Me.vwbunrui_Sheet1.Columns.Get(0).CellType = TextCellType1
        Me.vwbunrui_Sheet1.Columns.Get(0).Label = "料金分類" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "コード"
        Me.vwbunrui_Sheet1.Columns.Get(0).Locked = False
        Me.vwbunrui_Sheet1.Columns.Get(0).Width = 63.0!
        ComboBoxCellType1.AllowEditorVerticalAlign = True
        ComboBoxCellType1.ButtonAlign = FarPoint.Win.ButtonAlign.Right
        Me.vwbunrui_Sheet1.Columns.Get(1).CellType = ComboBoxCellType1
        Me.vwbunrui_Sheet1.Columns.Get(1).Label = "施設"
        Me.vwbunrui_Sheet1.Columns.Get(1).Width = 169.0!
        Me.vwbunrui_Sheet1.Columns.Get(2).CellType = TextCellType2
        Me.vwbunrui_Sheet1.Columns.Get(2).Label = "料金分類名"
        Me.vwbunrui_Sheet1.Columns.Get(2).Width = 252.0!
        Me.vwbunrui_Sheet1.Columns.Get(3).CellType = TextCellType3
        Me.vwbunrui_Sheet1.Columns.Get(3).Label = "並び順"
        Me.vwbunrui_Sheet1.Columns.Get(4).CellType = CheckBoxCellType1
        Me.vwbunrui_Sheet1.Columns.Get(4).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwbunrui_Sheet1.Columns.Get(4).Label = "無効"
        Me.vwbunrui_Sheet1.Columns.Get(4).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwbunrui_Sheet1.Columns.Get(4).Width = 40.0!
        Me.vwbunrui_Sheet1.RestrictRows = True
        Me.vwbunrui_Sheet1.RowHeader.Columns.Default.Resizable = False
        Me.vwbunrui_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.A1
        '
        'vwryokin
        '
        Me.vwryokin.AccessibleDescription = "vwryokin, Sheet1, Row 0, Column 0, "
        Me.vwryokin.Location = New System.Drawing.Point(961, 234)
        Me.vwryokin.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.vwryokin.Name = "vwryokin"
        Me.vwryokin.Sheets.AddRange(New FarPoint.Win.Spread.SheetView() {Me.vwryokin_Sheet1})
        Me.vwryokin.Size = New System.Drawing.Size(969, 435)
        Me.vwryokin.TabIndex = 139
        '
        'vwryokin_Sheet1
        '
        Me.vwryokin_Sheet1.Reset()
        Me.vwryokin_Sheet1.SheetName = "Sheet1"
        'Formulas and custom names must be loaded with R1C1 reference style
        Me.vwryokin_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.R1C1
        Me.vwryokin_Sheet1.ColumnCount = 9
        Me.vwryokin_Sheet1.ColumnHeader.Cells.Get(0, 0).Value = "料金分類" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "コード"
        Me.vwryokin_Sheet1.ColumnHeader.Cells.Get(0, 1).Value = "料金" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "コード"
        Me.vwryokin_Sheet1.ColumnHeader.Cells.Get(0, 2).Value = "料金名"
        Me.vwryokin_Sheet1.ColumnHeader.Cells.Get(0, 3).Value = "時間貸し"
        Me.vwryokin_Sheet1.ColumnHeader.Cells.Get(0, 4).Value = "利用料金" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "（1日,Lockout）"
        Me.vwryokin_Sheet1.ColumnHeader.Cells.Get(0, 5).Value = "並び順"
        Me.vwryokin_Sheet1.ColumnHeader.Cells.Get(0, 6).Value = "無効"
        Me.vwryokin_Sheet1.ColumnHeader.Cells.Get(0, 7).Value = "施設区分"
        Me.vwryokin_Sheet1.ColumnHeader.Cells.Get(0, 8).Value = "更新区分"
        Me.vwryokin_Sheet1.ColumnHeader.Rows.Get(0).Height = 34.0!
        Me.vwryokin_Sheet1.Columns.Get(0).CellType = TextCellType4
        Me.vwryokin_Sheet1.Columns.Get(0).Label = "料金分類" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "コード"
        Me.vwryokin_Sheet1.Columns.Get(0).Locked = False
        Me.vwryokin_Sheet1.Columns.Get(1).CellType = TextCellType5
        Me.vwryokin_Sheet1.Columns.Get(1).Label = "料金" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "コード"
        Me.vwryokin_Sheet1.Columns.Get(1).Locked = False
        Me.vwryokin_Sheet1.Columns.Get(2).CellType = TextCellType6
        Me.vwryokin_Sheet1.Columns.Get(2).Label = "料金名"
        Me.vwryokin_Sheet1.Columns.Get(2).Width = 249.0!
        NumberCellType1.DecimalPlaces = 0
        NumberCellType1.FixedPoint = False
        NumberCellType1.MaximumValue = 99999999.0R
        NumberCellType1.MinimumValue = 0R
        NumberCellType1.ShowSeparator = True
        Me.vwryokin_Sheet1.Columns.Get(3).CellType = NumberCellType1
        Me.vwryokin_Sheet1.Columns.Get(3).Label = "時間貸し"
        Me.vwryokin_Sheet1.Columns.Get(3).Width = 101.0!
        NumberCellType2.DecimalPlaces = 0
        NumberCellType2.FixedPoint = False
        NumberCellType2.MaximumValue = 99999999.0R
        NumberCellType2.MinimumValue = 0R
        NumberCellType2.ShowSeparator = True
        Me.vwryokin_Sheet1.Columns.Get(4).CellType = NumberCellType2
        Me.vwryokin_Sheet1.Columns.Get(4).Label = "利用料金" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "（1日,Lockout）"
        Me.vwryokin_Sheet1.Columns.Get(4).Width = 101.0!
        Me.vwryokin_Sheet1.Columns.Get(5).CellType = TextCellType7
        Me.vwryokin_Sheet1.Columns.Get(5).Label = "並び順"
        Me.vwryokin_Sheet1.Columns.Get(6).CellType = CheckBoxCellType2
        Me.vwryokin_Sheet1.Columns.Get(6).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwryokin_Sheet1.Columns.Get(6).Label = "無効"
        Me.vwryokin_Sheet1.Columns.Get(6).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwryokin_Sheet1.Columns.Get(6).Width = 40.0!
        Me.vwryokin_Sheet1.Columns.Get(7).Label = "施設区分"
        Me.vwryokin_Sheet1.Columns.Get(7).Visible = False
        Me.vwryokin_Sheet1.Columns.Get(8).Label = "更新区分"
        Me.vwryokin_Sheet1.Columns.Get(8).Visible = False
        Me.vwryokin_Sheet1.RestrictRows = True
        Me.vwryokin_Sheet1.RowHeader.Columns.Default.Resizable = False
        Me.vwryokin_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.A1
        '
        'vwbairitu
        '
        Me.vwbairitu.AccessibleDescription = "vwbairitu, Sheet1, Row 0, Column 0, "
        Me.vwbairitu.Location = New System.Drawing.Point(59, 724)
        Me.vwbairitu.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.vwbairitu.Name = "vwbairitu"
        Me.vwbairitu.Sheets.AddRange(New FarPoint.Win.Spread.SheetView() {Me.vwbairitu_Sheet1})
        Me.vwbairitu.Size = New System.Drawing.Size(951, 368)
        Me.vwbairitu.TabIndex = 140
        '
        'vwbairitu_Sheet1
        '
        Me.vwbairitu_Sheet1.Reset()
        Me.vwbairitu_Sheet1.SheetName = "Sheet1"
        'Formulas and custom names must be loaded with R1C1 reference style
        Me.vwbairitu_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.R1C1
        Me.vwbairitu_Sheet1.ColumnCount = 6
        Me.vwbairitu_Sheet1.ColumnHeader.Cells.Get(0, 0).Value = "倍率" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "コード"
        Me.vwbairitu_Sheet1.ColumnHeader.Cells.Get(0, 1).Value = "施設"
        Me.vwbairitu_Sheet1.ColumnHeader.Cells.Get(0, 2).Value = "倍率名"
        Me.vwbairitu_Sheet1.ColumnHeader.Cells.Get(0, 3).Value = "倍率"
        Me.vwbairitu_Sheet1.ColumnHeader.Cells.Get(0, 4).Value = "並び順"
        Me.vwbairitu_Sheet1.ColumnHeader.Cells.Get(0, 5).Value = "無効"
        Me.vwbairitu_Sheet1.ColumnHeader.Rows.Get(0).Height = 32.0!
        Me.vwbairitu_Sheet1.Columns.Get(0).CellType = TextCellType8
        Me.vwbairitu_Sheet1.Columns.Get(0).Label = "倍率" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "コード"
        Me.vwbairitu_Sheet1.Columns.Get(0).Locked = False
        Me.vwbairitu_Sheet1.Columns.Get(0).Width = 56.0!
        ComboBoxCellType2.AllowEditorVerticalAlign = True
        ComboBoxCellType2.ButtonAlign = FarPoint.Win.ButtonAlign.Right
        Me.vwbairitu_Sheet1.Columns.Get(1).CellType = ComboBoxCellType2
        Me.vwbairitu_Sheet1.Columns.Get(1).Label = "施設"
        Me.vwbairitu_Sheet1.Columns.Get(1).Width = 165.0!
        Me.vwbairitu_Sheet1.Columns.Get(2).CellType = TextCellType9
        Me.vwbairitu_Sheet1.Columns.Get(2).Label = "倍率名"
        Me.vwbairitu_Sheet1.Columns.Get(2).Width = 278.0!
        Me.vwbairitu_Sheet1.Columns.Get(4).CellType = TextCellType10
        Me.vwbairitu_Sheet1.Columns.Get(4).Label = "並び順"
        Me.vwbairitu_Sheet1.Columns.Get(5).CellType = CheckBoxCellType3
        Me.vwbairitu_Sheet1.Columns.Get(5).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.vwbairitu_Sheet1.Columns.Get(5).Label = "無効"
        Me.vwbairitu_Sheet1.Columns.Get(5).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwbairitu_Sheet1.Columns.Get(5).Width = 39.0!
        Me.vwbairitu_Sheet1.RestrictRows = True
        Me.vwbairitu_Sheet1.RowHeader.Columns.Default.Resizable = False
        Me.vwbairitu_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.A1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label1.Location = New System.Drawing.Point(567, 125)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(22, 15)
        Me.Label1.TabIndex = 156
        Me.Label1.Text = "月"
        '
        'txtkikantomonth
        '
        Me.txtkikantomonth.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txtkikantomonth.Location = New System.Drawing.Point(511, 121)
        Me.txtkikantomonth.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtkikantomonth.Name = "txtkikantomonth"
        Me.txtkikantomonth.Size = New System.Drawing.Size(47, 22)
        Me.txtkikantomonth.TabIndex = 155
        Me.txtkikantomonth.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label4.Location = New System.Drawing.Point(480, 125)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(22, 15)
        Me.Label4.TabIndex = 154
        Me.Label4.Text = "年"
        '
        'txtkikantoyear
        '
        Me.txtkikantoyear.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txtkikantoyear.Location = New System.Drawing.Point(393, 121)
        Me.txtkikantoyear.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtkikantoyear.Name = "txtkikantoyear"
        Me.txtkikantoyear.Size = New System.Drawing.Size(77, 22)
        Me.txtkikantoyear.TabIndex = 153
        Me.txtkikantoyear.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.BackColor = System.Drawing.Color.Transparent
        Me.Label5.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label5.Location = New System.Drawing.Point(320, 125)
        Me.Label5.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(47, 15)
        Me.Label5.TabIndex = 152
        Me.Label5.Text = "月　～"
        '
        'txtkikanfrommonth
        '
        Me.txtkikanfrommonth.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txtkikanfrommonth.Location = New System.Drawing.Point(264, 121)
        Me.txtkikanfrommonth.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtkikanfrommonth.Name = "txtkikanfrommonth"
        Me.txtkikanfrommonth.Size = New System.Drawing.Size(47, 22)
        Me.txtkikanfrommonth.TabIndex = 151
        Me.txtkikanfrommonth.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.BackColor = System.Drawing.Color.Transparent
        Me.Label11.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label11.Location = New System.Drawing.Point(233, 125)
        Me.Label11.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(22, 15)
        Me.Label11.TabIndex = 150
        Me.Label11.Text = "年"
        '
        'txtkikanfromyear
        '
        Me.txtkikanfromyear.Font = New System.Drawing.Font("MS UI Gothic", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txtkikanfromyear.Location = New System.Drawing.Point(147, 121)
        Me.txtkikanfromyear.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtkikanfromyear.Name = "txtkikanfromyear"
        Me.txtkikanfromyear.Size = New System.Drawing.Size(77, 22)
        Me.txtkikanfromyear.TabIndex = 149
        Me.txtkikanfromyear.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'EXTM0104
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.BackColor = System.Drawing.Color.White
        Me.BackgroundImage = Global.EXTM.My.Resources.Resources.背景
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(1979, 1202)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtkikantomonth)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtkikantoyear)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.txtkikanfrommonth)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.txtkikanfromyear)
        Me.Controls.Add(Me.vwbairitu)
        Me.Controls.Add(Me.vwryokin)
        Me.Controls.Add(Me.vwbunrui)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.btnsinki)
        Me.Controls.Add(Me.Label34)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.btnback)
        Me.Controls.Add(Me.btninsert)
        Me.DoubleBuffered = True
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Name = "EXTM0104"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ちゃり～ん。　利用料マスタメンテ"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.vwbunrui, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.vwbunrui_Sheet1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.vwryokin, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.vwryokin_Sheet1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.vwbairitu, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.vwbairitu_Sheet1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnsinki As System.Windows.Forms.Button
    Friend WithEvents Label34 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents cmbkikan As System.Windows.Forms.ComboBox
    Friend WithEvents rdosumi As System.Windows.Forms.RadioButton
    Friend WithEvents rdoshinki As System.Windows.Forms.RadioButton
    Friend WithEvents btnback As System.Windows.Forms.Button
    Friend WithEvents btninsert As System.Windows.Forms.Button
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents vwbunrui As FarPoint.Win.Spread.FpSpread
    Friend WithEvents vwbunrui_Sheet1 As FarPoint.Win.Spread.SheetView
    Friend WithEvents vwryokin As FarPoint.Win.Spread.FpSpread
    Friend WithEvents vwryokin_Sheet1 As FarPoint.Win.Spread.SheetView
    Friend WithEvents vwbairitu As FarPoint.Win.Spread.FpSpread
    Friend WithEvents vwbairitu_Sheet1 As FarPoint.Win.Spread.SheetView
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtkikantomonth As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtkikantoyear As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtkikanfrommonth As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents txtkikanfromyear As System.Windows.Forms.TextBox
End Class
