<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class EXTY0102
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
        Dim ButtonCellType1 As FarPoint.Win.Spread.CellType.ButtonCellType = New FarPoint.Win.Spread.CellType.ButtonCellType()
        Dim TextCellType2 As FarPoint.Win.Spread.CellType.TextCellType = New FarPoint.Win.Spread.CellType.TextCellType()
        Dim TextCellType3 As FarPoint.Win.Spread.CellType.TextCellType = New FarPoint.Win.Spread.CellType.TextCellType()
        Dim TextCellType4 As FarPoint.Win.Spread.CellType.TextCellType = New FarPoint.Win.Spread.CellType.TextCellType()
        Dim TextCellType5 As FarPoint.Win.Spread.CellType.TextCellType = New FarPoint.Win.Spread.CellType.TextCellType()
        Dim NumberCellType1 As FarPoint.Win.Spread.CellType.NumberCellType = New FarPoint.Win.Spread.CellType.NumberCellType()
        Dim ButtonCellType2 As FarPoint.Win.Spread.CellType.ButtonCellType = New FarPoint.Win.Spread.CellType.ButtonCellType()
        Dim TextCellType6 As FarPoint.Win.Spread.CellType.TextCellType = New FarPoint.Win.Spread.CellType.TextCellType()
        Dim TextCellType7 As FarPoint.Win.Spread.CellType.TextCellType = New FarPoint.Win.Spread.CellType.TextCellType()
        Dim TextCellType8 As FarPoint.Win.Spread.CellType.TextCellType = New FarPoint.Win.Spread.CellType.TextCellType()
        Dim TextCellType9 As FarPoint.Win.Spread.CellType.TextCellType = New FarPoint.Win.Spread.CellType.TextCellType()
        Dim TextCellType10 As FarPoint.Win.Spread.CellType.TextCellType = New FarPoint.Win.Spread.CellType.TextCellType()
        Dim TextCellType11 As FarPoint.Win.Spread.CellType.TextCellType = New FarPoint.Win.Spread.CellType.TextCellType()
        Dim ComboBoxCellType2 As FarPoint.Win.Spread.CellType.ComboBoxCellType = New FarPoint.Win.Spread.CellType.ComboBoxCellType()
        Dim ButtonCellType3 As FarPoint.Win.Spread.CellType.ButtonCellType = New FarPoint.Win.Spread.CellType.ButtonCellType()
        Dim TextCellType12 As FarPoint.Win.Spread.CellType.TextCellType = New FarPoint.Win.Spread.CellType.TextCellType()
        Dim ComboBoxCellType3 As FarPoint.Win.Spread.CellType.ComboBoxCellType = New FarPoint.Win.Spread.CellType.ComboBoxCellType()
        Dim TextCellType13 As FarPoint.Win.Spread.CellType.TextCellType = New FarPoint.Win.Spread.CellType.TextCellType()
        Dim ComboBoxCellType4 As FarPoint.Win.Spread.CellType.ComboBoxCellType = New FarPoint.Win.Spread.CellType.ComboBoxCellType()
        Dim NumberCellType2 As FarPoint.Win.Spread.CellType.NumberCellType = New FarPoint.Win.Spread.CellType.NumberCellType()
        Dim ButtonCellType4 As FarPoint.Win.Spread.CellType.ButtonCellType = New FarPoint.Win.Spread.CellType.ButtonCellType()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(EXTY0102))
        Me.grpSearchCondition = New System.Windows.Forms.GroupBox()
        Me.btnErrDisp = New System.Windows.Forms.Button()
        Me.dtpDspTo = New Common.DateTimePickerEx()
        Me.dtpDspFrom = New Common.DateTimePickerEx()
        Me.btnDataDsp = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnRef = New System.Windows.Forms.Button()
        Me.txtFilePath = New System.Windows.Forms.TextBox()
        Me.Label32 = New System.Windows.Forms.Label()
        Me.btnCsvInsert = New System.Windows.Forms.Button()
        Me.rdoProcUpd = New System.Windows.Forms.RadioButton()
        Me.rdoProcNew = New System.Windows.Forms.RadioButton()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.btnAdd = New System.Windows.Forms.Button()
        Me.btnBack = New System.Windows.Forms.Button()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.vwAlsokData = New FarPoint.Win.Spread.FpSpread()
        Me.vwAlsokData_Sheet1 = New FarPoint.Win.Spread.SheetView()
        Me.vwDegitalCashData = New FarPoint.Win.Spread.FpSpread()
        Me.vwDegitalCashData_Sheet1 = New FarPoint.Win.Spread.SheetView()
        Me.grpSearchCondition.SuspendLayout()
        CType(Me.vwAlsokData, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.vwAlsokData_Sheet1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.vwDegitalCashData, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.vwDegitalCashData_Sheet1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'grpSearchCondition
        '
        Me.grpSearchCondition.BackColor = System.Drawing.Color.Transparent
        Me.grpSearchCondition.Controls.Add(Me.btnErrDisp)
        Me.grpSearchCondition.Controls.Add(Me.dtpDspTo)
        Me.grpSearchCondition.Controls.Add(Me.dtpDspFrom)
        Me.grpSearchCondition.Controls.Add(Me.btnDataDsp)
        Me.grpSearchCondition.Controls.Add(Me.Label2)
        Me.grpSearchCondition.Controls.Add(Me.Label1)
        Me.grpSearchCondition.Controls.Add(Me.btnRef)
        Me.grpSearchCondition.Controls.Add(Me.txtFilePath)
        Me.grpSearchCondition.Controls.Add(Me.Label32)
        Me.grpSearchCondition.Controls.Add(Me.btnCsvInsert)
        Me.grpSearchCondition.Controls.Add(Me.rdoProcUpd)
        Me.grpSearchCondition.Controls.Add(Me.rdoProcNew)
        Me.grpSearchCondition.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.grpSearchCondition.Location = New System.Drawing.Point(34, 12)
        Me.grpSearchCondition.Name = "grpSearchCondition"
        Me.grpSearchCondition.Size = New System.Drawing.Size(1061, 167)
        Me.grpSearchCondition.TabIndex = 128
        Me.grpSearchCondition.TabStop = False
        '
        'btnErrDisp
        '
        Me.btnErrDisp.Location = New System.Drawing.Point(945, 55)
        Me.btnErrDisp.Name = "btnErrDisp"
        Me.btnErrDisp.Size = New System.Drawing.Size(110, 21)
        Me.btnErrDisp.TabIndex = 142
        Me.btnErrDisp.Text = "エラー内容表示"
        Me.btnErrDisp.UseVisualStyleBackColor = True
        '
        'dtpDspTo
        '
        Me.dtpDspTo.Location = New System.Drawing.Point(364, 124)
        Me.dtpDspTo.Name = "dtpDspTo"
        Me.dtpDspTo.Size = New System.Drawing.Size(144, 25)
        Me.dtpDspTo.TabIndex = 141
        '
        'dtpDspFrom
        '
        Me.dtpDspFrom.Location = New System.Drawing.Point(189, 124)
        Me.dtpDspFrom.Name = "dtpDspFrom"
        Me.dtpDspFrom.Size = New System.Drawing.Size(143, 25)
        Me.dtpDspFrom.TabIndex = 140
        '
        'btnDataDsp
        '
        Me.btnDataDsp.Location = New System.Drawing.Point(514, 124)
        Me.btnDataDsp.Name = "btnDataDsp"
        Me.btnDataDsp.Size = New System.Drawing.Size(104, 21)
        Me.btnDataDsp.TabIndex = 139
        Me.btnDataDsp.Text = "表示"
        Me.btnDataDsp.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label2.Location = New System.Drawing.Point(338, 128)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(20, 13)
        Me.Label2.TabIndex = 136
        Me.Label2.Text = "～"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label1.Location = New System.Drawing.Point(72, 128)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(107, 13)
        Me.Label1.TabIndex = 133
        Me.Label1.Text = "現金機利用日　＊"
        '
        'btnRef
        '
        Me.btnRef.Location = New System.Drawing.Point(677, 55)
        Me.btnRef.Name = "btnRef"
        Me.btnRef.Size = New System.Drawing.Size(103, 21)
        Me.btnRef.TabIndex = 131
        Me.btnRef.Text = "ファイル選択"
        Me.btnRef.UseVisualStyleBackColor = True
        '
        'txtFilePath
        '
        Me.txtFilePath.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txtFilePath.Location = New System.Drawing.Point(203, 56)
        Me.txtFilePath.Name = "txtFilePath"
        Me.txtFilePath.Size = New System.Drawing.Size(468, 20)
        Me.txtFilePath.TabIndex = 130
        '
        'Label32
        '
        Me.Label32.AutoSize = True
        Me.Label32.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label32.Location = New System.Drawing.Point(72, 59)
        Me.Label32.Name = "Label32"
        Me.Label32.Size = New System.Drawing.Size(110, 13)
        Me.Label32.TabIndex = 129
        Me.Label32.Text = "PC上のCSVファイル"
        '
        'btnCsvInsert
        '
        Me.btnCsvInsert.Location = New System.Drawing.Point(786, 55)
        Me.btnCsvInsert.Name = "btnCsvInsert"
        Me.btnCsvInsert.Size = New System.Drawing.Size(104, 21)
        Me.btnCsvInsert.TabIndex = 75
        Me.btnCsvInsert.Text = "表示"
        Me.btnCsvInsert.UseVisualStyleBackColor = True
        '
        'rdoProcUpd
        '
        Me.rdoProcUpd.AutoSize = True
        Me.rdoProcUpd.Location = New System.Drawing.Point(40, 97)
        Me.rdoProcUpd.Name = "rdoProcUpd"
        Me.rdoProcUpd.Size = New System.Drawing.Size(190, 17)
        Me.rdoProcUpd.TabIndex = 32
        Me.rdoProcUpd.TabStop = True
        Me.rdoProcUpd.Text = "紐付け済みデータの表示、修正"
        Me.rdoProcUpd.UseVisualStyleBackColor = True
        '
        'rdoProcNew
        '
        Me.rdoProcNew.AutoSize = True
        Me.rdoProcNew.Location = New System.Drawing.Point(40, 25)
        Me.rdoProcNew.Name = "rdoProcNew"
        Me.rdoProcNew.Size = New System.Drawing.Size(77, 17)
        Me.rdoProcNew.TabIndex = 31
        Me.rdoProcNew.TabStop = True
        Me.rdoProcNew.Text = "新規取込"
        Me.rdoProcNew.UseVisualStyleBackColor = True
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label3.Location = New System.Drawing.Point(34, 660)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(117, 13)
        Me.Label3.TabIndex = 134
        Me.Label3.Text = "電子マネー入金登録"
        '
        'btnAdd
        '
        Me.btnAdd.Location = New System.Drawing.Point(37, 829)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(78, 29)
        Me.btnAdd.TabIndex = 140
        Me.btnAdd.Text = "行追加"
        Me.btnAdd.UseVisualStyleBackColor = True
        '
        'btnBack
        '
        Me.btnBack.Location = New System.Drawing.Point(383, 912)
        Me.btnBack.Name = "btnBack"
        Me.btnBack.Size = New System.Drawing.Size(145, 38)
        Me.btnBack.TabIndex = 141
        Me.btnBack.Text = "戻る"
        Me.btnBack.UseVisualStyleBackColor = True
        '
        'btnUpdate
        '
        Me.btnUpdate.Location = New System.Drawing.Point(979, 912)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(145, 38)
        Me.btnUpdate.TabIndex = 143
        Me.btnUpdate.Text = "登録"
        Me.btnUpdate.UseVisualStyleBackColor = True
        '
        'vwAlsokData
        '
        Me.vwAlsokData.AccessibleDescription = "vwAlsokData, Sheet1, Row 0, Column 0, "
        Me.vwAlsokData.Location = New System.Drawing.Point(34, 185)
        Me.vwAlsokData.Name = "vwAlsokData"
        Me.vwAlsokData.Sheets.AddRange(New FarPoint.Win.Spread.SheetView() {Me.vwAlsokData_Sheet1})
        Me.vwAlsokData.Size = New System.Drawing.Size(1201, 439)
        Me.vwAlsokData.TabIndex = 144
        '
        'vwAlsokData_Sheet1
        '
        Me.vwAlsokData_Sheet1.Reset()
        Me.vwAlsokData_Sheet1.SheetName = "Sheet1"
        'Formulas and custom names must be loaded with R1C1 reference style
        Me.vwAlsokData_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.R1C1
        Me.vwAlsokData_Sheet1.ColumnCount = 15
        Me.vwAlsokData_Sheet1.ColumnHeader.Cells.Get(0, 0).Value = "日付"
        Me.vwAlsokData_Sheet1.ColumnHeader.Cells.Get(0, 1).Value = "催事名／アーティスト"
        Me.vwAlsokData_Sheet1.ColumnHeader.Cells.Get(0, 2).Value = "　"
        Me.vwAlsokData_Sheet1.ColumnHeader.Cells.Get(0, 3).Value = "レジNO"
        Me.vwAlsokData_Sheet1.ColumnHeader.Cells.Get(0, 4).Value = "レジ名"
        Me.vwAlsokData_Sheet1.ColumnHeader.Cells.Get(0, 5).Value = "店舗NO"
        Me.vwAlsokData_Sheet1.ColumnHeader.Cells.Get(0, 6).Value = "店舗名"
        Me.vwAlsokData_Sheet1.ColumnHeader.Cells.Get(0, 7).Value = "投入金額"
        Me.vwAlsokData_Sheet1.ColumnHeader.Cells.Get(0, 8).Value = "日毎合算"
        Me.vwAlsokData_Sheet1.ColumnHeader.Cells.Get(0, 9).Value = "　"
        Me.vwAlsokData_Sheet1.ColumnHeader.Cells.Get(0, 10).Value = "入金機"
        Me.vwAlsokData_Sheet1.ColumnHeader.Cells.Get(0, 11).Value = "連番"
        Me.vwAlsokData_Sheet1.ColumnHeader.Cells.Get(0, 12).Value = "施設区分"
        Me.vwAlsokData_Sheet1.ColumnHeader.Cells.Get(0, 13).Value = "削除区分"
        Me.vwAlsokData_Sheet1.ColumnHeader.Cells.Get(0, 14).Value = "契約先"
        TextCellType1.AllowEditorVerticalAlign = True
        Me.vwAlsokData_Sheet1.Columns.Get(0).CellType = TextCellType1
        Me.vwAlsokData_Sheet1.Columns.Get(0).Label = "日付"
        Me.vwAlsokData_Sheet1.Columns.Get(0).Locked = True
        Me.vwAlsokData_Sheet1.Columns.Get(0).Width = 75.0!
        ComboBoxCellType1.AllowEditorVerticalAlign = True
        ComboBoxCellType1.ButtonAlign = FarPoint.Win.ButtonAlign.Right
        Me.vwAlsokData_Sheet1.Columns.Get(1).CellType = ComboBoxCellType1
        Me.vwAlsokData_Sheet1.Columns.Get(1).Label = "催事名／アーティスト"
        Me.vwAlsokData_Sheet1.Columns.Get(1).Width = 202.0!
        ButtonCellType1.ButtonColor2 = System.Drawing.SystemColors.ButtonFace
        ButtonCellType1.Text = "選択"
        Me.vwAlsokData_Sheet1.Columns.Get(2).CellType = ButtonCellType1
        Me.vwAlsokData_Sheet1.Columns.Get(2).Label = "　"
        TextCellType2.AllowEditorVerticalAlign = True
        Me.vwAlsokData_Sheet1.Columns.Get(3).CellType = TextCellType2
        Me.vwAlsokData_Sheet1.Columns.Get(3).Label = "レジNO"
        Me.vwAlsokData_Sheet1.Columns.Get(3).Locked = True
        TextCellType3.AllowEditorVerticalAlign = True
        Me.vwAlsokData_Sheet1.Columns.Get(4).CellType = TextCellType3
        Me.vwAlsokData_Sheet1.Columns.Get(4).Label = "レジ名"
        Me.vwAlsokData_Sheet1.Columns.Get(4).Locked = True
        Me.vwAlsokData_Sheet1.Columns.Get(4).Width = 206.0!
        TextCellType4.AllowEditorVerticalAlign = True
        Me.vwAlsokData_Sheet1.Columns.Get(5).CellType = TextCellType4
        Me.vwAlsokData_Sheet1.Columns.Get(5).Label = "店舗NO"
        Me.vwAlsokData_Sheet1.Columns.Get(5).Locked = True
        TextCellType5.AllowEditorVerticalAlign = True
        Me.vwAlsokData_Sheet1.Columns.Get(6).CellType = TextCellType5
        Me.vwAlsokData_Sheet1.Columns.Get(6).Label = "店舗名"
        Me.vwAlsokData_Sheet1.Columns.Get(6).Locked = True
        Me.vwAlsokData_Sheet1.Columns.Get(6).Width = 204.0!
        NumberCellType1.AllowEditorVerticalAlign = True
        NumberCellType1.DecimalPlaces = 0
        NumberCellType1.FixedPoint = False
        NumberCellType1.MaximumValue = 99999999.0R
        NumberCellType1.MinimumValue = -10000000.0R
        NumberCellType1.ReadOnly = True
        NumberCellType1.Separator = ","
        NumberCellType1.ShowSeparator = True
        Me.vwAlsokData_Sheet1.Columns.Get(7).CellType = NumberCellType1
        Me.vwAlsokData_Sheet1.Columns.Get(7).Label = "投入金額"
        Me.vwAlsokData_Sheet1.Columns.Get(7).Locked = True
        Me.vwAlsokData_Sheet1.Columns.Get(7).Width = 122.0!
        Me.vwAlsokData_Sheet1.Columns.Get(8).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.General
        Me.vwAlsokData_Sheet1.Columns.Get(8).Label = "日毎合算"
        Me.vwAlsokData_Sheet1.Columns.Get(8).Locked = True
        Me.vwAlsokData_Sheet1.Columns.Get(8).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.vwAlsokData_Sheet1.Columns.Get(8).Width = 94.0!
        ButtonCellType2.ButtonColor2 = System.Drawing.SystemColors.ButtonFace
        ButtonCellType2.Text = "削除"
        Me.vwAlsokData_Sheet1.Columns.Get(9).CellType = ButtonCellType2
        Me.vwAlsokData_Sheet1.Columns.Get(9).Label = "　"
        Me.vwAlsokData_Sheet1.Columns.Get(9).Locked = False
        Me.vwAlsokData_Sheet1.Columns.Get(9).Width = 59.0!
        TextCellType6.AllowEditorVerticalAlign = True
        Me.vwAlsokData_Sheet1.Columns.Get(10).CellType = TextCellType6
        Me.vwAlsokData_Sheet1.Columns.Get(10).Label = "入金機"
        Me.vwAlsokData_Sheet1.Columns.Get(10).Visible = False
        TextCellType7.AllowEditorVerticalAlign = True
        Me.vwAlsokData_Sheet1.Columns.Get(11).CellType = TextCellType7
        Me.vwAlsokData_Sheet1.Columns.Get(11).Label = "連番"
        Me.vwAlsokData_Sheet1.Columns.Get(11).Visible = False
        TextCellType8.AllowEditorVerticalAlign = True
        Me.vwAlsokData_Sheet1.Columns.Get(12).CellType = TextCellType8
        Me.vwAlsokData_Sheet1.Columns.Get(12).Label = "施設区分"
        Me.vwAlsokData_Sheet1.Columns.Get(12).Visible = False
        TextCellType9.AllowEditorVerticalAlign = True
        Me.vwAlsokData_Sheet1.Columns.Get(13).CellType = TextCellType9
        Me.vwAlsokData_Sheet1.Columns.Get(13).Label = "削除区分"
        Me.vwAlsokData_Sheet1.Columns.Get(13).Visible = False
        TextCellType10.AllowEditorVerticalAlign = True
        Me.vwAlsokData_Sheet1.Columns.Get(14).CellType = TextCellType10
        Me.vwAlsokData_Sheet1.Columns.Get(14).Label = "契約先"
        Me.vwAlsokData_Sheet1.Columns.Get(14).Visible = False
        Me.vwAlsokData_Sheet1.RowHeader.Columns.Default.Resizable = False
        Me.vwAlsokData_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.A1
        '
        'vwDegitalCashData
        '
        Me.vwDegitalCashData.AccessibleDescription = "vwDegitalCashData, Sheet1, Row 0, Column 0, "
        Me.vwDegitalCashData.Location = New System.Drawing.Point(37, 676)
        Me.vwDegitalCashData.Name = "vwDegitalCashData"
        Me.vwDegitalCashData.Sheets.AddRange(New FarPoint.Win.Spread.SheetView() {Me.vwDegitalCashData_Sheet1})
        Me.vwDegitalCashData.Size = New System.Drawing.Size(1110, 143)
        Me.vwDegitalCashData.TabIndex = 149
        '
        'vwDegitalCashData_Sheet1
        '
        Me.vwDegitalCashData_Sheet1.Reset()
        Me.vwDegitalCashData_Sheet1.SheetName = "Sheet1"
        'Formulas and custom names must be loaded with R1C1 reference style
        Me.vwDegitalCashData_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.R1C1
        Me.vwDegitalCashData_Sheet1.ColumnCount = 13
        Me.vwDegitalCashData_Sheet1.ColumnHeader.Cells.Get(0, 0).Value = "日付"
        Me.vwDegitalCashData_Sheet1.ColumnHeader.Cells.Get(0, 1).Value = "催事名／アーティスト"
        Me.vwDegitalCashData_Sheet1.ColumnHeader.Cells.Get(0, 2).Value = "　"
        Me.vwDegitalCashData_Sheet1.ColumnHeader.Cells.Get(0, 3).Value = "レジNO"
        Me.vwDegitalCashData_Sheet1.ColumnHeader.Cells.Get(0, 4).Value = "レジ名"
        Me.vwDegitalCashData_Sheet1.ColumnHeader.Cells.Get(0, 5).Value = "店舗NO"
        Me.vwDegitalCashData_Sheet1.ColumnHeader.Cells.Get(0, 6).Value = "店舗名"
        Me.vwDegitalCashData_Sheet1.ColumnHeader.Cells.Get(0, 7).Value = "投入金額"
        Me.vwDegitalCashData_Sheet1.ColumnHeader.Cells.Get(0, 8).Value = "　"
        Me.vwDegitalCashData_Sheet1.ColumnHeader.Cells.Get(0, 9).Value = "入金機番号"
        Me.vwDegitalCashData_Sheet1.ColumnHeader.Cells.Get(0, 10).Value = "連番"
        Me.vwDegitalCashData_Sheet1.ColumnHeader.Cells.Get(0, 11).Value = "施設区分"
        Me.vwDegitalCashData_Sheet1.ColumnHeader.Cells.Get(0, 12).Value = "削除区分"
        TextCellType11.AllowEditorVerticalAlign = True
        Me.vwDegitalCashData_Sheet1.Columns.Get(0).CellType = TextCellType11
        Me.vwDegitalCashData_Sheet1.Columns.Get(0).Label = "日付"
        Me.vwDegitalCashData_Sheet1.Columns.Get(0).Width = 75.0!
        ComboBoxCellType2.AllowEditorVerticalAlign = True
        ComboBoxCellType2.ButtonAlign = FarPoint.Win.ButtonAlign.Right
        Me.vwDegitalCashData_Sheet1.Columns.Get(1).CellType = ComboBoxCellType2
        Me.vwDegitalCashData_Sheet1.Columns.Get(1).Label = "催事名／アーティスト"
        Me.vwDegitalCashData_Sheet1.Columns.Get(1).Width = 202.0!
        ButtonCellType3.ButtonColor2 = System.Drawing.SystemColors.ButtonFace
        ButtonCellType3.Text = "選択"
        Me.vwDegitalCashData_Sheet1.Columns.Get(2).CellType = ButtonCellType3
        Me.vwDegitalCashData_Sheet1.Columns.Get(2).Label = "　"
        TextCellType12.AllowEditorVerticalAlign = True
        Me.vwDegitalCashData_Sheet1.Columns.Get(3).CellType = TextCellType12
        Me.vwDegitalCashData_Sheet1.Columns.Get(3).Label = "レジNO"
        Me.vwDegitalCashData_Sheet1.Columns.Get(3).Locked = True
        ComboBoxCellType3.AllowEditorVerticalAlign = True
        ComboBoxCellType3.ButtonAlign = FarPoint.Win.ButtonAlign.Right
        Me.vwDegitalCashData_Sheet1.Columns.Get(4).CellType = ComboBoxCellType3
        Me.vwDegitalCashData_Sheet1.Columns.Get(4).Label = "レジ名"
        Me.vwDegitalCashData_Sheet1.Columns.Get(4).Width = 210.0!
        TextCellType13.AllowEditorVerticalAlign = True
        Me.vwDegitalCashData_Sheet1.Columns.Get(5).CellType = TextCellType13
        Me.vwDegitalCashData_Sheet1.Columns.Get(5).Label = "店舗NO"
        Me.vwDegitalCashData_Sheet1.Columns.Get(5).Locked = True
        ComboBoxCellType4.AllowEditorVerticalAlign = True
        ComboBoxCellType4.ButtonAlign = FarPoint.Win.ButtonAlign.Right
        Me.vwDegitalCashData_Sheet1.Columns.Get(6).CellType = ComboBoxCellType4
        Me.vwDegitalCashData_Sheet1.Columns.Get(6).Label = "店舗名"
        Me.vwDegitalCashData_Sheet1.Columns.Get(6).Width = 204.0!
        NumberCellType2.AllowEditorVerticalAlign = True
        NumberCellType2.DecimalPlaces = 0
        NumberCellType2.FixedPoint = False
        NumberCellType2.MaximumValue = 99999999.0R
        NumberCellType2.MinimumValue = 0.0R
        NumberCellType2.Separator = ","
        NumberCellType2.ShowSeparator = True
        Me.vwDegitalCashData_Sheet1.Columns.Get(7).CellType = NumberCellType2
        Me.vwDegitalCashData_Sheet1.Columns.Get(7).Label = "投入金額"
        Me.vwDegitalCashData_Sheet1.Columns.Get(7).Width = 122.0!
        ButtonCellType4.ButtonColor2 = System.Drawing.SystemColors.ButtonFace
        ButtonCellType4.Text = "削除"
        Me.vwDegitalCashData_Sheet1.Columns.Get(8).CellType = ButtonCellType4
        Me.vwDegitalCashData_Sheet1.Columns.Get(8).Label = "　"
        Me.vwDegitalCashData_Sheet1.Columns.Get(9).Label = "入金機番号"
        Me.vwDegitalCashData_Sheet1.Columns.Get(9).Visible = False
        Me.vwDegitalCashData_Sheet1.Columns.Get(9).Width = 77.0!
        Me.vwDegitalCashData_Sheet1.Columns.Get(10).Label = "連番"
        Me.vwDegitalCashData_Sheet1.Columns.Get(10).Visible = False
        Me.vwDegitalCashData_Sheet1.Columns.Get(11).Label = "施設区分"
        Me.vwDegitalCashData_Sheet1.Columns.Get(11).Visible = False
        Me.vwDegitalCashData_Sheet1.Columns.Get(12).Label = "削除区分"
        Me.vwDegitalCashData_Sheet1.Columns.Get(12).Visible = False
        Me.vwDegitalCashData_Sheet1.RowHeader.Columns.Default.Resizable = False
        Me.vwDegitalCashData_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.A1
        '
        'EXTY0102
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.BackColor = System.Drawing.Color.White
        Me.BackgroundImage = Global.EXTY.My.Resources.Resources.背景
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(1362, 966)
        Me.Controls.Add(Me.vwDegitalCashData)
        Me.Controls.Add(Me.vwAlsokData)
        Me.Controls.Add(Me.btnUpdate)
        Me.Controls.Add(Me.btnBack)
        Me.Controls.Add(Me.btnAdd)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.grpSearchCondition)
        Me.DoubleBuffered = True
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "EXTY0102"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ちゃり～ん。　ALSOK・電子マネー入金登録"
        Me.grpSearchCondition.ResumeLayout(False)
        Me.grpSearchCondition.PerformLayout()
        CType(Me.vwAlsokData, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.vwAlsokData_Sheet1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.vwDegitalCashData, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.vwDegitalCashData_Sheet1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents grpSearchCondition As System.Windows.Forms.GroupBox
    Friend WithEvents btnCsvInsert As System.Windows.Forms.Button
    Friend WithEvents rdoProcUpd As System.Windows.Forms.RadioButton
    Friend WithEvents rdoProcNew As System.Windows.Forms.RadioButton
    Friend WithEvents Label32 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnRef As System.Windows.Forms.Button
    Friend WithEvents txtFilePath As System.Windows.Forms.TextBox
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents btnDataDsp As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Friend WithEvents btnBack As System.Windows.Forms.Button
    Friend WithEvents btnUpdate As System.Windows.Forms.Button
    Friend WithEvents vwAlsokData As FarPoint.Win.Spread.FpSpread
    Friend WithEvents vwAlsokData_Sheet1 As FarPoint.Win.Spread.SheetView
    Friend WithEvents vwDegitalCashData As FarPoint.Win.Spread.FpSpread
    Friend WithEvents vwDegitalCashData_Sheet1 As FarPoint.Win.Spread.SheetView
    Friend WithEvents dtpDspTo As Common.DateTimePickerEx
    Friend WithEvents dtpDspFrom As Common.DateTimePickerEx
    Friend WithEvents btnErrDisp As System.Windows.Forms.Button
End Class
