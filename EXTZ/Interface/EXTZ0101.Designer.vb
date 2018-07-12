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
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Label66 = New System.Windows.Forms.Label()
        Me.Label65 = New System.Windows.Forms.Label()
        Me.Label64 = New System.Windows.Forms.Label()
        Me.TextBox54 = New System.Windows.Forms.TextBox()
        Me.Label63 = New System.Windows.Forms.Label()
        Me.TextBox53 = New System.Windows.Forms.TextBox()
        Me.Label62 = New System.Windows.Forms.Label()
        Me.TextBox52 = New System.Windows.Forms.TextBox()
        Me.Label54 = New System.Windows.Forms.Label()
        Me.Label53 = New System.Windows.Forms.Label()
        Me.TextBox47 = New System.Windows.Forms.TextBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.RadioButton2 = New System.Windows.Forms.RadioButton()
        Me.RadioButton1 = New System.Windows.Forms.RadioButton()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.DateTimePickerEx2 = New Common.DateTimePickerEx()
        Me.DateTimePickerEx1 = New Common.DateTimePickerEx()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.FpSpread1 = New FarPoint.Win.Spread.FpSpread()
        Me.FpSpread1_Sheet1 = New FarPoint.Win.Spread.SheetView()
        Me.FpSpread1_Sheet2 = New FarPoint.Win.Spread.SheetView()
        Me.Button6 = New System.Windows.Forms.Button()
        Me.Panel1.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        CType(Me.FpSpread1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.FpSpread1_Sheet1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.FpSpread1_Sheet2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(56, 944)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(175, 36)
        Me.Button5.TabIndex = 22
        Me.Button5.Text = "戻る"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'Label66
        '
        Me.Label66.AutoSize = True
        Me.Label66.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label66.Location = New System.Drawing.Point(21, 34)
        Me.Label66.Name = "Label66"
        Me.Label66.Size = New System.Drawing.Size(59, 13)
        Me.Label66.TabIndex = 23
        Me.Label66.Text = "検索対象"
        '
        'Label65
        '
        Me.Label65.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label65.Location = New System.Drawing.Point(21, 66)
        Me.Label65.Name = "Label65"
        Me.Label65.Size = New System.Drawing.Size(102, 42)
        Me.Label65.TabIndex = 25
        Me.Label65.Text = "利用日"
        '
        'Label64
        '
        Me.Label64.AutoSize = True
        Me.Label64.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label64.Location = New System.Drawing.Point(21, 99)
        Me.Label64.Name = "Label64"
        Me.Label64.Size = New System.Drawing.Size(82, 13)
        Me.Label64.TabIndex = 27
        Me.Label64.Text = "利用者名(ｶﾅ)" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'TextBox54
        '
        Me.TextBox54.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.TextBox54.Location = New System.Drawing.Point(224, 96)
        Me.TextBox54.Name = "TextBox54"
        Me.TextBox54.Size = New System.Drawing.Size(306, 20)
        Me.TextBox54.TabIndex = 28
        '
        'Label63
        '
        Me.Label63.AutoSize = True
        Me.Label63.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label63.Location = New System.Drawing.Point(21, 137)
        Me.Label63.Name = "Label63"
        Me.Label63.Size = New System.Drawing.Size(129, 13)
        Me.Label63.TabIndex = 29
        Me.Label63.Text = "催事名／アーティスト名"
        '
        'TextBox53
        '
        Me.TextBox53.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.TextBox53.Location = New System.Drawing.Point(224, 134)
        Me.TextBox53.Name = "TextBox53"
        Me.TextBox53.Size = New System.Drawing.Size(306, 20)
        Me.TextBox53.TabIndex = 30
        '
        'Label62
        '
        Me.Label62.AutoSize = True
        Me.Label62.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label62.Location = New System.Drawing.Point(21, 171)
        Me.Label62.Name = "Label62"
        Me.Label62.Size = New System.Drawing.Size(59, 13)
        Me.Label62.TabIndex = 31
        Me.Label62.Text = "予約番号"
        '
        'TextBox52
        '
        Me.TextBox52.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.TextBox52.Location = New System.Drawing.Point(224, 168)
        Me.TextBox52.Name = "TextBox52"
        Me.TextBox52.Size = New System.Drawing.Size(112, 20)
        Me.TextBox52.TabIndex = 32
        '
        'Label54
        '
        Me.Label54.AutoSize = True
        Me.Label54.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label54.Location = New System.Drawing.Point(652, 65)
        Me.Label54.Name = "Label54"
        Me.Label54.Size = New System.Drawing.Size(59, 13)
        Me.Label54.TabIndex = 43
        Me.Label54.Text = "利用者名"
        '
        'Label53
        '
        Me.Label53.AutoSize = True
        Me.Label53.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label53.Location = New System.Drawing.Point(316, 66)
        Me.Label53.Name = "Label53"
        Me.Label53.Size = New System.Drawing.Size(20, 13)
        Me.Label53.TabIndex = 46
        Me.Label53.Text = "～"
        '
        'TextBox47
        '
        Me.TextBox47.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.TextBox47.Location = New System.Drawing.Point(751, 63)
        Me.TextBox47.Name = "TextBox47"
        Me.TextBox47.Size = New System.Drawing.Size(364, 20)
        Me.TextBox47.TabIndex = 47
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.RadioButton2)
        Me.Panel1.Controls.Add(Me.RadioButton1)
        Me.Panel1.Location = New System.Drawing.Point(153, 30)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(222, 22)
        Me.Panel1.TabIndex = 28
        '
        'RadioButton2
        '
        Me.RadioButton2.AutoSize = True
        Me.RadioButton2.Location = New System.Drawing.Point(99, 3)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(65, 17)
        Me.RadioButton2.TabIndex = 30
        Me.RadioButton2.Text = "スタジオ"
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'RadioButton1
        '
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.Checked = True
        Me.RadioButton1.Location = New System.Drawing.Point(3, 3)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(66, 17)
        Me.RadioButton1.TabIndex = 29
        Me.RadioButton1.TabStop = True
        Me.RadioButton1.Text = "シアター"
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Checked = True
        Me.CheckBox1.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox1.Location = New System.Drawing.Point(24, 204)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(161, 17)
        Me.CheckBox1.TabIndex = 32
        Me.CheckBox1.Text = "未完了の予約のみを表示"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.BackgroundImage = Global.EXTZ.My.Resources.Resources.マスタ検索
        Me.Button2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Button2.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Button2.Location = New System.Drawing.Point(536, 92)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(36, 27)
        Me.Button2.TabIndex = 83
        Me.Button2.UseVisualStyleBackColor = True
        '
        'GroupBox5
        '
        Me.GroupBox5.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox5.Controls.Add(Me.DateTimePickerEx2)
        Me.GroupBox5.Controls.Add(Me.DateTimePickerEx1)
        Me.GroupBox5.Controls.Add(Me.Button4)
        Me.GroupBox5.Controls.Add(Me.Button2)
        Me.GroupBox5.Controls.Add(Me.CheckBox1)
        Me.GroupBox5.Controls.Add(Me.Panel1)
        Me.GroupBox5.Controls.Add(Me.TextBox47)
        Me.GroupBox5.Controls.Add(Me.Label53)
        Me.GroupBox5.Controls.Add(Me.Label54)
        Me.GroupBox5.Controls.Add(Me.TextBox52)
        Me.GroupBox5.Controls.Add(Me.Label62)
        Me.GroupBox5.Controls.Add(Me.TextBox53)
        Me.GroupBox5.Controls.Add(Me.Label63)
        Me.GroupBox5.Controls.Add(Me.TextBox54)
        Me.GroupBox5.Controls.Add(Me.Label64)
        Me.GroupBox5.Controls.Add(Me.Label65)
        Me.GroupBox5.Controls.Add(Me.Label66)
        Me.GroupBox5.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.GroupBox5.Location = New System.Drawing.Point(12, 42)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(1539, 234)
        Me.GroupBox5.TabIndex = 77
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "検索条件"
        '
        'DateTimePickerEx2
        '
        Me.DateTimePickerEx2.Location = New System.Drawing.Point(342, 58)
        Me.DateTimePickerEx2.Name = "DateTimePickerEx2"
        Me.DateTimePickerEx2.Size = New System.Drawing.Size(140, 25)
        Me.DateTimePickerEx2.TabIndex = 88
        '
        'DateTimePickerEx1
        '
        Me.DateTimePickerEx1.Location = New System.Drawing.Point(167, 60)
        Me.DateTimePickerEx1.Name = "DateTimePickerEx1"
        Me.DateTimePickerEx1.Size = New System.Drawing.Size(143, 23)
        Me.DateTimePickerEx1.TabIndex = 87
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(1133, 152)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(96, 23)
        Me.Button4.TabIndex = 86
        Me.Button4.Text = "検索"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'FpSpread1
        '
        Me.FpSpread1.AccessibleDescription = "FpSpread1, Sheet2, Row 0, Column 0, "
        Me.FpSpread1.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Never
        Me.FpSpread1.Location = New System.Drawing.Point(12, 282)
        Me.FpSpread1.Name = "FpSpread1"
        Me.FpSpread1.Sheets.AddRange(New FarPoint.Win.Spread.SheetView() {Me.FpSpread1_Sheet1, Me.FpSpread1_Sheet2})
        Me.FpSpread1.Size = New System.Drawing.Size(1620, 616)
        Me.FpSpread1.TabIndex = 90
        Me.FpSpread1.TabStripPolicy = FarPoint.Win.Spread.TabStripPolicy.Never
        Me.FpSpread1.VerticalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.AsNeeded
        '
        'FpSpread1_Sheet1
        '
        Me.FpSpread1_Sheet1.Reset()
        Me.FpSpread1_Sheet1.SheetName = "Sheet1"
        'Formulas and custom names must be loaded with R1C1 reference style
        Me.FpSpread1_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.R1C1
        Me.FpSpread1_Sheet1.ColumnCount = 16
        Me.FpSpread1_Sheet1.RowCount = 30
        Me.FpSpread1_Sheet1.ActiveColumnIndex = 1
        Me.FpSpread1_Sheet1.Cells.Get(0, 1).Value = "仮未"
        Me.FpSpread1_Sheet1.Cells.Get(1, 1).Value = "仮"
        Me.FpSpread1_Sheet1.Cells.Get(2, 1).Value = "正式"
        Me.FpSpread1_Sheet1.Cells.Get(3, 1).Value = "完了"
        Me.FpSpread1_Sheet1.ColumnHeader.Cells.Get(0, 0).Value = "選択"
        Me.FpSpread1_Sheet1.ColumnHeader.Cells.Get(0, 1).Value = "ｽﾃｰﾀｽ"
        Me.FpSpread1_Sheet1.ColumnHeader.Cells.Get(0, 2).Value = "予約番号"
        Me.FpSpread1_Sheet1.ColumnHeader.Cells.Get(0, 3).Value = "利用開始日"
        Me.FpSpread1_Sheet1.ColumnHeader.Cells.Get(0, 4).Value = "利用終了日"
        Me.FpSpread1_Sheet1.ColumnHeader.Cells.Get(0, 5).Value = "開始時間"
        Me.FpSpread1_Sheet1.ColumnHeader.Cells.Get(0, 6).Value = "終了時間"
        Me.FpSpread1_Sheet1.ColumnHeader.Cells.Get(0, 7).Value = "催事名"
        Me.FpSpread1_Sheet1.ColumnHeader.Cells.Get(0, 8).Value = "利用形状"
        Me.FpSpread1_Sheet1.ColumnHeader.Cells.Get(0, 9).Value = "利用者名"
        Me.FpSpread1_Sheet1.ColumnHeader.Cells.Get(0, 10).Value = "責任者名"
        Me.FpSpread1_Sheet1.ColumnHeader.Cells.Get(0, 11).Value = "承認人数"
        Me.FpSpread1_Sheet1.ColumnHeader.Cells.Get(0, 12).Value = "利用料"
        Me.FpSpread1_Sheet1.ColumnHeader.Cells.Get(0, 13).Value = "設備入力ｽﾃｰﾀｽ"
        Me.FpSpread1_Sheet1.ColumnHeader.Cells.Get(0, 14).Value = "付帯設備"
        Me.FpSpread1_Sheet1.ColumnHeader.Cells.Get(0, 15).Value = "　"
        Me.FpSpread1_Sheet1.ColumnHeader.Rows.Get(0).Height = 32.0!
        Me.FpSpread1_Sheet1.Columns.Get(0).CellType = CheckBoxCellType1
        Me.FpSpread1_Sheet1.Columns.Get(0).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.FpSpread1_Sheet1.Columns.Get(0).Label = "選択"
        Me.FpSpread1_Sheet1.Columns.Get(0).Width = 35.0!
        Me.FpSpread1_Sheet1.Columns.Get(1).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.FpSpread1_Sheet1.Columns.Get(1).Label = "ｽﾃｰﾀｽ"
        Me.FpSpread1_Sheet1.Columns.Get(1).Width = 49.0!
        Me.FpSpread1_Sheet1.Columns.Get(2).Label = "予約番号"
        Me.FpSpread1_Sheet1.Columns.Get(2).Width = 104.0!
        Me.FpSpread1_Sheet1.Columns.Get(3).Label = "利用開始日"
        Me.FpSpread1_Sheet1.Columns.Get(3).Width = 103.0!
        Me.FpSpread1_Sheet1.Columns.Get(4).Label = "利用終了日"
        Me.FpSpread1_Sheet1.Columns.Get(4).Width = 103.0!
        Me.FpSpread1_Sheet1.Columns.Get(5).Label = "開始時間"
        Me.FpSpread1_Sheet1.Columns.Get(5).Width = 81.0!
        Me.FpSpread1_Sheet1.Columns.Get(6).Label = "終了時間"
        Me.FpSpread1_Sheet1.Columns.Get(6).Width = 81.0!
        Me.FpSpread1_Sheet1.Columns.Get(7).Label = "催事名"
        Me.FpSpread1_Sheet1.Columns.Get(7).Width = 261.0!
        Me.FpSpread1_Sheet1.Columns.Get(8).Label = "利用形状"
        Me.FpSpread1_Sheet1.Columns.Get(8).Width = 84.0!
        Me.FpSpread1_Sheet1.Columns.Get(9).Label = "利用者名"
        Me.FpSpread1_Sheet1.Columns.Get(9).Width = 157.0!
        Me.FpSpread1_Sheet1.Columns.Get(10).Label = "責任者名"
        Me.FpSpread1_Sheet1.Columns.Get(10).Width = 157.0!
        Me.FpSpread1_Sheet1.Columns.Get(11).Label = "承認人数"
        Me.FpSpread1_Sheet1.Columns.Get(11).Width = 68.0!
        ButtonCellType1.ButtonColor2 = System.Drawing.SystemColors.ButtonFace
        ButtonCellType1.Text = "確認・編集"
        Me.FpSpread1_Sheet1.Columns.Get(15).CellType = ButtonCellType1
        Me.FpSpread1_Sheet1.Columns.Get(15).Label = "　"
        Me.FpSpread1_Sheet1.Columns.Get(15).Width = 101.0!
        Me.FpSpread1_Sheet1.RowHeader.Columns.Default.Resizable = False
        Me.FpSpread1_Sheet1.Rows.Get(0).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet1.Rows.Get(1).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet1.Rows.Get(2).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet1.Rows.Get(3).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet1.Rows.Get(4).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet1.Rows.Get(5).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet1.Rows.Get(6).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet1.Rows.Get(7).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet1.Rows.Get(8).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet1.Rows.Get(9).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet1.Rows.Get(10).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet1.Rows.Get(11).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet1.Rows.Get(12).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet1.Rows.Get(13).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet1.Rows.Get(14).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet1.Rows.Get(15).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet1.Rows.Get(16).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet1.Rows.Get(17).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet1.Rows.Get(18).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet1.Rows.Get(19).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet1.Rows.Get(20).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet1.Rows.Get(21).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet1.Rows.Get(22).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet1.Rows.Get(23).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet1.Rows.Get(24).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet1.Rows.Get(25).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet1.Rows.Get(26).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet1.Rows.Get(27).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet1.Rows.Get(28).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet1.Rows.Get(29).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.A1
        '
        'FpSpread1_Sheet2
        '
        Me.FpSpread1_Sheet2.Reset()
        Me.FpSpread1_Sheet2.SheetName = "Sheet2"
        'Formulas and custom names must be loaded with R1C1 reference style
        Me.FpSpread1_Sheet2.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.R1C1
        Me.FpSpread1_Sheet2.ColumnCount = 16
        Me.FpSpread1_Sheet2.RowCount = 30
        Me.FpSpread1_Sheet2.Cells.Get(0, 1).Value = "仮未"
        Me.FpSpread1_Sheet2.Cells.Get(1, 1).Value = "仮"
        Me.FpSpread1_Sheet2.Cells.Get(2, 1).Value = "正式"
        Me.FpSpread1_Sheet2.Cells.Get(3, 1).Value = "完了"
        Me.FpSpread1_Sheet2.ColumnHeader.Cells.Get(0, 0).Value = "選択"
        Me.FpSpread1_Sheet2.ColumnHeader.Cells.Get(0, 1).Value = "ｽﾃｰﾀｽ"
        Me.FpSpread1_Sheet2.ColumnHeader.Cells.Get(0, 2).Value = "予約番号"
        Me.FpSpread1_Sheet2.ColumnHeader.Cells.Get(0, 3).Value = "スタジオ"
        Me.FpSpread1_Sheet2.ColumnHeader.Cells.Get(0, 4).Value = "利用開始日"
        Me.FpSpread1_Sheet2.ColumnHeader.Cells.Get(0, 5).Value = "利用終了日"
        Me.FpSpread1_Sheet2.ColumnHeader.Cells.Get(0, 6).Value = "開始時間"
        Me.FpSpread1_Sheet2.ColumnHeader.Cells.Get(0, 7).Value = "終了時間"
        Me.FpSpread1_Sheet2.ColumnHeader.Cells.Get(0, 8).Value = "アーティスト名"
        Me.FpSpread1_Sheet2.ColumnHeader.Cells.Get(0, 9).Value = "利用者名"
        Me.FpSpread1_Sheet2.ColumnHeader.Cells.Get(0, 10).Value = "責任者名"
        Me.FpSpread1_Sheet2.ColumnHeader.Cells.Get(0, 11).Value = "承認人数"
        Me.FpSpread1_Sheet2.ColumnHeader.Cells.Get(0, 12).Value = "利用料"
        Me.FpSpread1_Sheet2.ColumnHeader.Cells.Get(0, 13).Value = "設備入力ｽﾃｰﾀｽ"
        Me.FpSpread1_Sheet2.ColumnHeader.Cells.Get(0, 14).Value = "付帯設備"
        Me.FpSpread1_Sheet2.ColumnHeader.Cells.Get(0, 15).Value = "　"
        Me.FpSpread1_Sheet2.ColumnHeader.Rows.Get(0).Height = 30.0!
        Me.FpSpread1_Sheet2.Columns.Get(0).CellType = CheckBoxCellType2
        Me.FpSpread1_Sheet2.Columns.Get(0).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.FpSpread1_Sheet2.Columns.Get(0).Label = "選択"
        Me.FpSpread1_Sheet2.Columns.Get(0).Width = 35.0!
        Me.FpSpread1_Sheet2.Columns.Get(1).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.FpSpread1_Sheet2.Columns.Get(1).Label = "ｽﾃｰﾀｽ"
        Me.FpSpread1_Sheet2.Columns.Get(1).Width = 46.0!
        Me.FpSpread1_Sheet2.Columns.Get(2).Label = "予約番号"
        Me.FpSpread1_Sheet2.Columns.Get(2).Width = 104.0!
        Me.FpSpread1_Sheet2.Columns.Get(4).Label = "利用開始日"
        Me.FpSpread1_Sheet2.Columns.Get(4).Width = 103.0!
        Me.FpSpread1_Sheet2.Columns.Get(5).Label = "利用終了日"
        Me.FpSpread1_Sheet2.Columns.Get(5).Width = 103.0!
        Me.FpSpread1_Sheet2.Columns.Get(6).Label = "開始時間"
        Me.FpSpread1_Sheet2.Columns.Get(6).Width = 81.0!
        Me.FpSpread1_Sheet2.Columns.Get(7).Label = "終了時間"
        Me.FpSpread1_Sheet2.Columns.Get(7).Width = 81.0!
        Me.FpSpread1_Sheet2.Columns.Get(8).Label = "アーティスト名"
        Me.FpSpread1_Sheet2.Columns.Get(8).Width = 218.0!
        Me.FpSpread1_Sheet2.Columns.Get(9).Label = "利用者名"
        Me.FpSpread1_Sheet2.Columns.Get(9).Width = 157.0!
        Me.FpSpread1_Sheet2.Columns.Get(10).Label = "責任者名"
        Me.FpSpread1_Sheet2.Columns.Get(10).Width = 157.0!
        Me.FpSpread1_Sheet2.Columns.Get(11).Label = "承認人数"
        Me.FpSpread1_Sheet2.Columns.Get(11).Width = 68.0!
        ButtonCellType2.ButtonColor2 = System.Drawing.SystemColors.ButtonFace
        ButtonCellType2.Text = "確認・編集"
        Me.FpSpread1_Sheet2.Columns.Get(15).CellType = ButtonCellType2
        Me.FpSpread1_Sheet2.Columns.Get(15).Label = "　"
        Me.FpSpread1_Sheet2.Columns.Get(15).Width = 101.0!
        Me.FpSpread1_Sheet2.RowHeader.Columns.Default.Resizable = False
        Me.FpSpread1_Sheet2.Rows.Get(0).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet2.Rows.Get(1).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet2.Rows.Get(2).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet2.Rows.Get(3).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet2.Rows.Get(4).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet2.Rows.Get(5).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet2.Rows.Get(6).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet2.Rows.Get(7).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet2.Rows.Get(8).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet2.Rows.Get(9).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet2.Rows.Get(10).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet2.Rows.Get(11).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet2.Rows.Get(12).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet2.Rows.Get(13).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet2.Rows.Get(14).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet2.Rows.Get(15).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet2.Rows.Get(16).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet2.Rows.Get(17).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet2.Rows.Get(18).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet2.Rows.Get(19).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet2.Rows.Get(20).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet2.Rows.Get(21).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet2.Rows.Get(22).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet2.Rows.Get(23).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet2.Rows.Get(24).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet2.Rows.Get(25).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet2.Rows.Get(26).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet2.Rows.Get(27).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet2.Rows.Get(28).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet2.Rows.Get(29).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.FpSpread1_Sheet2.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.A1
        '
        'Button6
        '
        Me.Button6.Location = New System.Drawing.Point(1376, 944)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(175, 36)
        Me.Button6.TabIndex = 81
        Me.Button6.Text = "選択確定"
        Me.Button6.UseVisualStyleBackColor = True
        '
        'EXTZ0101
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.BackgroundImage = Global.EXTZ.My.Resources.Resources.背景
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(1664, 1042)
        Me.Controls.Add(Me.FpSpread1)
        Me.Controls.Add(Me.Button6)
        Me.Controls.Add(Me.GroupBox5)
        Me.Controls.Add(Me.Button5)
        Me.DoubleBuffered = True
        Me.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "EXTZ0101"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ちゃり～ん。　予約一覧"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox5.PerformLayout()
        CType(Me.FpSpread1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.FpSpread1_Sheet1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.FpSpread1_Sheet2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents Label66 As System.Windows.Forms.Label
    Friend WithEvents Label65 As System.Windows.Forms.Label
    Friend WithEvents Label64 As System.Windows.Forms.Label
    Friend WithEvents TextBox54 As System.Windows.Forms.TextBox
    Friend WithEvents Label63 As System.Windows.Forms.Label
    Friend WithEvents TextBox53 As System.Windows.Forms.TextBox
    Friend WithEvents Label62 As System.Windows.Forms.Label
    Friend WithEvents TextBox52 As System.Windows.Forms.TextBox
    Friend WithEvents Label54 As System.Windows.Forms.Label
    Friend WithEvents Label53 As System.Windows.Forms.Label
    Friend WithEvents TextBox47 As System.Windows.Forms.TextBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents RadioButton2 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton1 As System.Windows.Forms.RadioButton
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Button6 As System.Windows.Forms.Button
    Friend WithEvents FpSpread1 As FarPoint.Win.Spread.FpSpread
    Friend WithEvents FpSpread1_Sheet1 As FarPoint.Win.Spread.SheetView
    Friend WithEvents DateTimePickerEx2 As Common.DateTimePickerEx
    Friend WithEvents DateTimePickerEx1 As Common.DateTimePickerEx
    Friend WithEvents FpSpread1_Sheet2 As FarPoint.Win.Spread.SheetView

End Class
