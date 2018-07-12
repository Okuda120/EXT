<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class EXTZ0204
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
        Dim NumberCellType1 As FarPoint.Win.Spread.CellType.NumberCellType = New FarPoint.Win.Spread.CellType.NumberCellType()
        Dim NumberCellType2 As FarPoint.Win.Spread.CellType.NumberCellType = New FarPoint.Win.Spread.CellType.NumberCellType()
        Dim ButtonCellType1 As FarPoint.Win.Spread.CellType.ButtonCellType = New FarPoint.Win.Spread.CellType.ButtonCellType()
        Dim NumberCellType3 As FarPoint.Win.Spread.CellType.NumberCellType = New FarPoint.Win.Spread.CellType.NumberCellType()
        Dim NumberCellType4 As FarPoint.Win.Spread.CellType.NumberCellType = New FarPoint.Win.Spread.CellType.NumberCellType()
        Dim ButtonCellType2 As FarPoint.Win.Spread.CellType.ButtonCellType = New FarPoint.Win.Spread.CellType.ButtonCellType()
        Dim ButtonCellType3 As FarPoint.Win.Spread.CellType.ButtonCellType = New FarPoint.Win.Spread.CellType.ButtonCellType()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(EXTZ0204))
        Me.Label65 = New System.Windows.Forms.Label()
        Me.btnBack = New System.Windows.Forms.Button()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.lblTaxSagaku = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.lblSagakuKin = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.btnComplate = New System.Windows.Forms.Button()
        Me.fbBillPay = New FarPoint.Win.Spread.FpSpread()
        Me.fbBillPay_Sheet1 = New FarPoint.Win.Spread.SheetView()
        Me.fbRiyoryo = New FarPoint.Win.Spread.FpSpread()
        Me.fbRiyoryo_Sheet1 = New FarPoint.Win.Spread.SheetView()
        Me.fbFutai = New FarPoint.Win.Spread.FpSpread()
        Me.fbFutai_Sheet1 = New FarPoint.Win.Spread.SheetView()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtTitle1 = New System.Windows.Forms.TextBox()
        Me.txtTitle2 = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        CType(Me.fbBillPay, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fbBillPay_Sheet1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fbRiyoryo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fbRiyoryo_Sheet1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fbFutai, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fbFutai_Sheet1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label65
        '
        Me.Label65.BackColor = System.Drawing.Color.Transparent
        Me.Label65.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label65.Location = New System.Drawing.Point(12, 18)
        Me.Label65.Name = "Label65"
        Me.Label65.Size = New System.Drawing.Size(103, 21)
        Me.Label65.TabIndex = 25
        Me.Label65.Text = "■お客様請求額"
        '
        'btnBack
        '
        Me.btnBack.Location = New System.Drawing.Point(658, 641)
        Me.btnBack.Name = "btnBack"
        Me.btnBack.Size = New System.Drawing.Size(96, 23)
        Me.btnBack.TabIndex = 89
        Me.btnBack.Text = "戻る"
        Me.btnBack.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label4.Location = New System.Drawing.Point(12, 225)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(72, 21)
        Me.Label4.TabIndex = 93
        Me.Label4.Text = "利用料"
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.Transparent
        Me.Label5.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label5.Location = New System.Drawing.Point(12, 125)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(103, 21)
        Me.Label5.TabIndex = 94
        Me.Label5.Text = "■売上計上設定"
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.Color.Transparent
        Me.Label10.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label10.Location = New System.Drawing.Point(12, 372)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(103, 21)
        Me.Label10.TabIndex = 104
        Me.Label10.Text = "付帯設備"
        '
        'lblTaxSagaku
        '
        Me.lblTaxSagaku.BackColor = System.Drawing.Color.Silver
        Me.lblTaxSagaku.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTaxSagaku.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.lblTaxSagaku.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblTaxSagaku.Location = New System.Drawing.Point(121, 597)
        Me.lblTaxSagaku.Margin = New System.Windows.Forms.Padding(3)
        Me.lblTaxSagaku.Name = "lblTaxSagaku"
        Me.lblTaxSagaku.Size = New System.Drawing.Size(88, 20)
        Me.lblTaxSagaku.TabIndex = 112
        Me.lblTaxSagaku.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label13
        '
        Me.Label13.BackColor = System.Drawing.Color.Transparent
        Me.Label13.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label13.Location = New System.Drawing.Point(12, 600)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(103, 17)
        Me.Label13.TabIndex = 111
        Me.Label13.Text = "消費税差額"
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblSagakuKin
        '
        Me.lblSagakuKin.BackColor = System.Drawing.Color.Silver
        Me.lblSagakuKin.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblSagakuKin.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.lblSagakuKin.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblSagakuKin.Location = New System.Drawing.Point(121, 565)
        Me.lblSagakuKin.Margin = New System.Windows.Forms.Padding(3)
        Me.lblSagakuKin.Name = "lblSagakuKin"
        Me.lblSagakuKin.Size = New System.Drawing.Size(88, 20)
        Me.lblSagakuKin.TabIndex = 110
        Me.lblSagakuKin.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label15
        '
        Me.Label15.BackColor = System.Drawing.Color.Transparent
        Me.Label15.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label15.Location = New System.Drawing.Point(12, 569)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(105, 21)
        Me.Label15.TabIndex = 109
        Me.Label15.Text = "請求金額との差額"
        '
        'btnComplate
        '
        Me.btnComplate.Location = New System.Drawing.Point(881, 641)
        Me.btnComplate.Name = "btnComplate"
        Me.btnComplate.Size = New System.Drawing.Size(140, 23)
        Me.btnComplate.TabIndex = 86
        Me.btnComplate.Text = "入力完了"
        Me.btnComplate.UseVisualStyleBackColor = True
        '
        'fbBillPay
        '
        Me.fbBillPay.AccessibleDescription = "fbBillPay, Sheet1"
        Me.fbBillPay.Location = New System.Drawing.Point(15, 42)
        Me.fbBillPay.Name = "fbBillPay"
        Me.fbBillPay.Sheets.AddRange(New FarPoint.Win.Spread.SheetView() {Me.fbBillPay_Sheet1})
        Me.fbBillPay.Size = New System.Drawing.Size(763, 63)
        Me.fbBillPay.TabIndex = 113
        '
        'fbBillPay_Sheet1
        '
        Me.fbBillPay_Sheet1.Reset()
        Me.fbBillPay_Sheet1.SheetName = "Sheet1"
        'Formulas and custom names must be loaded with R1C1 reference style
        Me.fbBillPay_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.R1C1
        Me.fbBillPay_Sheet1.ColumnCount = 15
        Me.fbBillPay_Sheet1.RowCount = 0
        Me.fbBillPay_Sheet1.ActiveColumnIndex = -1
        Me.fbBillPay_Sheet1.ActiveRowIndex = -1
        Me.fbBillPay_Sheet1.ColumnHeader.Cells.Get(0, 0).Value = "請求内容"
        Me.fbBillPay_Sheet1.ColumnHeader.Cells.Get(0, 1).Value = "確定額"
        Me.fbBillPay_Sheet1.ColumnHeader.Cells.Get(0, 2).Value = "グロス調整"
        Me.fbBillPay_Sheet1.ColumnHeader.Cells.Get(0, 3).Value = "小計（税抜）"
        Me.fbBillPay_Sheet1.ColumnHeader.Cells.Get(0, 4).Value = "消費税割合"
        Me.fbBillPay_Sheet1.ColumnHeader.Cells.Get(0, 5).Value = "消費税"
        Me.fbBillPay_Sheet1.ColumnHeader.Cells.Get(0, 6).Value = "請求金額"
        Me.fbBillPay_Sheet1.Columns.Get(0).Label = "請求内容"
        Me.fbBillPay_Sheet1.Columns.Get(0).Width = 142.0!
        Me.fbBillPay_Sheet1.Columns.Get(1).Label = "確定額"
        Me.fbBillPay_Sheet1.Columns.Get(1).Width = 92.0!
        Me.fbBillPay_Sheet1.Columns.Get(2).Label = "グロス調整"
        Me.fbBillPay_Sheet1.Columns.Get(2).Width = 92.0!
        Me.fbBillPay_Sheet1.Columns.Get(3).Label = "小計（税抜）"
        Me.fbBillPay_Sheet1.Columns.Get(3).Width = 92.0!
        Me.fbBillPay_Sheet1.Columns.Get(4).Label = "消費税割合"
        Me.fbBillPay_Sheet1.Columns.Get(4).Width = 88.0!
        Me.fbBillPay_Sheet1.Columns.Get(5).Label = "消費税"
        Me.fbBillPay_Sheet1.Columns.Get(5).Width = 100.0!
        Me.fbBillPay_Sheet1.Columns.Get(6).Label = "請求金額"
        Me.fbBillPay_Sheet1.Columns.Get(6).Width = 100.0!
        Me.fbBillPay_Sheet1.RowHeader.Columns.Default.Resizable = False
        Me.fbBillPay_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.A1
        '
        'fbRiyoryo
        '
        Me.fbRiyoryo.AccessibleDescription = "FpSpread2, Sheet1, Row 0, Column 0, 2015年5月"
        Me.fbRiyoryo.Location = New System.Drawing.Point(15, 249)
        Me.fbRiyoryo.Name = "fbRiyoryo"
        Me.fbRiyoryo.Sheets.AddRange(New FarPoint.Win.Spread.SheetView() {Me.fbRiyoryo_Sheet1})
        Me.fbRiyoryo.Size = New System.Drawing.Size(1554, 97)
        Me.fbRiyoryo.TabIndex = 114
        Me.fbRiyoryo.SetViewportLeftColumn(0, 0, 2)
        '
        'fbRiyoryo_Sheet1
        '
        Me.fbRiyoryo_Sheet1.Reset()
        Me.fbRiyoryo_Sheet1.SheetName = "Sheet1"
        'Formulas and custom names must be loaded with R1C1 reference style
        Me.fbRiyoryo_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.R1C1
        Me.fbRiyoryo_Sheet1.ColumnCount = 30
        Me.fbRiyoryo_Sheet1.RowCount = 0
        Me.fbRiyoryo_Sheet1.ActiveColumnIndex = -1
        Me.fbRiyoryo_Sheet1.ActiveRowIndex = -1
        Me.fbRiyoryo_Sheet1.ColumnHeader.Cells.Get(0, 0).Value = "利用年月" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "（発生月）"
        Me.fbRiyoryo_Sheet1.ColumnHeader.Cells.Get(0, 1).Value = "勘定科目名"
        Me.fbRiyoryo_Sheet1.ColumnHeader.Cells.Get(0, 2).Value = "細目名"
        Me.fbRiyoryo_Sheet1.ColumnHeader.Cells.Get(0, 3).Value = "内訳名"
        Me.fbRiyoryo_Sheet1.ColumnHeader.Cells.Get(0, 4).Value = "詳細名"
        Me.fbRiyoryo_Sheet1.ColumnHeader.Cells.Get(0, 5).Value = "利用料計上額"
        Me.fbRiyoryo_Sheet1.ColumnHeader.Cells.Get(0, 6).Value = "消費税"
        Me.fbRiyoryo_Sheet1.ColumnHeader.Cells.Get(0, 7).Value = "入力摘要1"
        Me.fbRiyoryo_Sheet1.ColumnHeader.Cells.Get(0, 8).Value = "入力摘要2"
        Me.fbRiyoryo_Sheet1.ColumnHeader.Cells.Get(0, 9).Value = "プロジェクト"
        Me.fbRiyoryo_Sheet1.ColumnHeader.Cells.Get(0, 10).Value = "プロジェクト" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "内訳"
        Me.fbRiyoryo_Sheet1.ColumnHeader.Cells.Get(0, 11).Value = "プロジェクト" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "一覧"
        Me.fbRiyoryo_Sheet1.ColumnHeader.Rows.Get(0).Height = 36.0!
        Me.fbRiyoryo_Sheet1.Columns.Get(0).BackColor = System.Drawing.Color.FromArgb(CType(CType(246, Byte), Integer), CType(CType(234, Byte), Integer), CType(CType(210, Byte), Integer))
        Me.fbRiyoryo_Sheet1.Columns.Get(0).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.fbRiyoryo_Sheet1.Columns.Get(0).Label = "利用年月" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "（発生月）"
        Me.fbRiyoryo_Sheet1.Columns.Get(0).Width = 78.0!
        Me.fbRiyoryo_Sheet1.Columns.Get(1).BackColor = System.Drawing.Color.FromArgb(CType(CType(246, Byte), Integer), CType(CType(234, Byte), Integer), CType(CType(210, Byte), Integer))
        Me.fbRiyoryo_Sheet1.Columns.Get(1).Label = "勘定科目名"
        Me.fbRiyoryo_Sheet1.Columns.Get(1).Width = 110.0!
        Me.fbRiyoryo_Sheet1.Columns.Get(2).BackColor = System.Drawing.Color.FromArgb(CType(CType(246, Byte), Integer), CType(CType(234, Byte), Integer), CType(CType(210, Byte), Integer))
        Me.fbRiyoryo_Sheet1.Columns.Get(2).Label = "細目名"
        Me.fbRiyoryo_Sheet1.Columns.Get(2).Width = 110.0!
        Me.fbRiyoryo_Sheet1.Columns.Get(3).BackColor = System.Drawing.Color.FromArgb(CType(CType(246, Byte), Integer), CType(CType(234, Byte), Integer), CType(CType(210, Byte), Integer))
        Me.fbRiyoryo_Sheet1.Columns.Get(3).Label = "内訳名"
        Me.fbRiyoryo_Sheet1.Columns.Get(3).Width = 110.0!
        Me.fbRiyoryo_Sheet1.Columns.Get(4).BackColor = System.Drawing.Color.FromArgb(CType(CType(246, Byte), Integer), CType(CType(234, Byte), Integer), CType(CType(210, Byte), Integer))
        Me.fbRiyoryo_Sheet1.Columns.Get(4).Label = "詳細名"
        Me.fbRiyoryo_Sheet1.Columns.Get(4).Width = 110.0!
        NumberCellType1.AllowEditorVerticalAlign = True
        NumberCellType1.DecimalPlaces = 0
        NumberCellType1.FixedPoint = False
        NumberCellType1.MaximumValue = 99999999999.0R
        NumberCellType1.MinimumValue = -99999999999.0R
        NumberCellType1.ShowSeparator = True
        Me.fbRiyoryo_Sheet1.Columns.Get(5).CellType = NumberCellType1
        Me.fbRiyoryo_Sheet1.Columns.Get(5).Label = "利用料計上額"
        Me.fbRiyoryo_Sheet1.Columns.Get(5).Width = 100.0!
        NumberCellType2.AllowEditorVerticalAlign = True
        NumberCellType2.DecimalPlaces = 0
        NumberCellType2.FixedPoint = False
        NumberCellType2.MaximumValue = 9999999999.0R
        NumberCellType2.MinimumValue = -9999999999.0R
        NumberCellType2.ShowSeparator = True
        Me.fbRiyoryo_Sheet1.Columns.Get(6).CellType = NumberCellType2
        Me.fbRiyoryo_Sheet1.Columns.Get(6).Label = "消費税"
        Me.fbRiyoryo_Sheet1.Columns.Get(6).Width = 100.0!
        Me.fbRiyoryo_Sheet1.Columns.Get(7).Label = "入力摘要1"
        Me.fbRiyoryo_Sheet1.Columns.Get(7).Width = 227.0!
        Me.fbRiyoryo_Sheet1.Columns.Get(8).Label = "入力摘要2"
        Me.fbRiyoryo_Sheet1.Columns.Get(8).Width = 227.0!
        Me.fbRiyoryo_Sheet1.Columns.Get(9).BackColor = System.Drawing.Color.FromArgb(CType(CType(246, Byte), Integer), CType(CType(234, Byte), Integer), CType(CType(210, Byte), Integer))
        Me.fbRiyoryo_Sheet1.Columns.Get(9).Label = "プロジェクト"
        Me.fbRiyoryo_Sheet1.Columns.Get(9).Width = 99.0!
        Me.fbRiyoryo_Sheet1.Columns.Get(10).BackColor = System.Drawing.Color.FromArgb(CType(CType(246, Byte), Integer), CType(CType(234, Byte), Integer), CType(CType(210, Byte), Integer))
        Me.fbRiyoryo_Sheet1.Columns.Get(10).Label = "プロジェクト" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "内訳"
        Me.fbRiyoryo_Sheet1.Columns.Get(10).Width = 156.0!
        ButtonCellType1.ButtonColor2 = System.Drawing.SystemColors.ButtonFace
        ButtonCellType1.Text = "表示"
        Me.fbRiyoryo_Sheet1.Columns.Get(11).CellType = ButtonCellType1
        Me.fbRiyoryo_Sheet1.Columns.Get(11).Label = "プロジェクト" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "一覧"
        Me.fbRiyoryo_Sheet1.Columns.Get(11).Width = 70.0!
        Me.fbRiyoryo_Sheet1.RowHeader.Columns.Default.Resizable = False
        Me.fbRiyoryo_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.A1
        '
        'fbFutai
        '
        Me.fbFutai.AccessibleDescription = "FpSpread3, Sheet1, Row 0, Column 0, 2015年5月"
        Me.fbFutai.Location = New System.Drawing.Point(12, 396)
        Me.fbFutai.Name = "fbFutai"
        Me.fbFutai.Sheets.AddRange(New FarPoint.Win.Spread.SheetView() {Me.fbFutai_Sheet1})
        Me.fbFutai.Size = New System.Drawing.Size(1591, 143)
        Me.fbFutai.TabIndex = 115
        Me.fbFutai.SetViewportLeftColumn(0, 0, 4)
        '
        'fbFutai_Sheet1
        '
        Me.fbFutai_Sheet1.Reset()
        Me.fbFutai_Sheet1.SheetName = "Sheet1"
        'Formulas and custom names must be loaded with R1C1 reference style
        Me.fbFutai_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.R1C1
        Me.fbFutai_Sheet1.ColumnCount = 30
        Me.fbFutai_Sheet1.RowCount = 0
        Me.fbFutai_Sheet1.ActiveColumnIndex = -1
        Me.fbFutai_Sheet1.ActiveRowIndex = -1
        Me.fbFutai_Sheet1.ColumnHeader.Cells.Get(0, 0).Value = "利用年月"
        Me.fbFutai_Sheet1.ColumnHeader.Cells.Get(0, 1).Value = "集計キー"
        Me.fbFutai_Sheet1.ColumnHeader.Cells.Get(0, 2).Value = "勘定科目名"
        Me.fbFutai_Sheet1.ColumnHeader.Cells.Get(0, 3).Value = "細目名"
        Me.fbFutai_Sheet1.ColumnHeader.Cells.Get(0, 4).Value = "内訳名"
        Me.fbFutai_Sheet1.ColumnHeader.Cells.Get(0, 5).Value = "詳細名"
        Me.fbFutai_Sheet1.ColumnHeader.Cells.Get(0, 6).Value = "付帯設備" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "計上額"
        Me.fbFutai_Sheet1.ColumnHeader.Cells.Get(0, 7).Value = "消費税"
        Me.fbFutai_Sheet1.ColumnHeader.Cells.Get(0, 8).Value = "入力摘要1"
        Me.fbFutai_Sheet1.ColumnHeader.Cells.Get(0, 9).Value = "入力摘要2"
        Me.fbFutai_Sheet1.ColumnHeader.Cells.Get(0, 10).Value = "プロジェクト"
        Me.fbFutai_Sheet1.ColumnHeader.Cells.Get(0, 11).Value = "プロジェクト" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "内訳"
        Me.fbFutai_Sheet1.ColumnHeader.Cells.Get(0, 12).Value = "コピー"
        Me.fbFutai_Sheet1.ColumnHeader.Cells.Get(0, 13).Value = "プロジェクト" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "一覧"
        Me.fbFutai_Sheet1.ColumnHeader.Rows.Get(0).Height = 41.0!
        Me.fbFutai_Sheet1.Columns.Get(0).BackColor = System.Drawing.Color.FromArgb(CType(CType(246, Byte), Integer), CType(CType(234, Byte), Integer), CType(CType(210, Byte), Integer))
        Me.fbFutai_Sheet1.Columns.Get(0).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.fbFutai_Sheet1.Columns.Get(0).Label = "利用年月"
        Me.fbFutai_Sheet1.Columns.Get(0).Width = 71.0!
        Me.fbFutai_Sheet1.Columns.Get(1).BackColor = System.Drawing.Color.FromArgb(CType(CType(246, Byte), Integer), CType(CType(234, Byte), Integer), CType(CType(210, Byte), Integer))
        Me.fbFutai_Sheet1.Columns.Get(1).Label = "集計キー"
        Me.fbFutai_Sheet1.Columns.Get(1).Width = 136.0!
        Me.fbFutai_Sheet1.Columns.Get(2).BackColor = System.Drawing.Color.FromArgb(CType(CType(246, Byte), Integer), CType(CType(234, Byte), Integer), CType(CType(210, Byte), Integer))
        Me.fbFutai_Sheet1.Columns.Get(2).Label = "勘定科目名"
        Me.fbFutai_Sheet1.Columns.Get(2).Width = 108.0!
        Me.fbFutai_Sheet1.Columns.Get(3).BackColor = System.Drawing.Color.FromArgb(CType(CType(246, Byte), Integer), CType(CType(234, Byte), Integer), CType(CType(210, Byte), Integer))
        Me.fbFutai_Sheet1.Columns.Get(3).Label = "細目名"
        Me.fbFutai_Sheet1.Columns.Get(3).Width = 108.0!
        Me.fbFutai_Sheet1.Columns.Get(4).BackColor = System.Drawing.Color.FromArgb(CType(CType(246, Byte), Integer), CType(CType(234, Byte), Integer), CType(CType(210, Byte), Integer))
        Me.fbFutai_Sheet1.Columns.Get(4).Label = "内訳名"
        Me.fbFutai_Sheet1.Columns.Get(4).Width = 108.0!
        Me.fbFutai_Sheet1.Columns.Get(5).BackColor = System.Drawing.Color.FromArgb(CType(CType(246, Byte), Integer), CType(CType(234, Byte), Integer), CType(CType(210, Byte), Integer))
        Me.fbFutai_Sheet1.Columns.Get(5).Label = "詳細名"
        Me.fbFutai_Sheet1.Columns.Get(5).Width = 108.0!
        NumberCellType3.AllowEditorVerticalAlign = True
        NumberCellType3.DecimalPlaces = 0
        NumberCellType3.FixedPoint = False
        NumberCellType3.MaximumValue = 99999999999.0R
        NumberCellType3.MinimumValue = -99999999999.0R
        NumberCellType3.ShowSeparator = True
        Me.fbFutai_Sheet1.Columns.Get(6).CellType = NumberCellType3
        Me.fbFutai_Sheet1.Columns.Get(6).Label = "付帯設備" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "計上額"
        Me.fbFutai_Sheet1.Columns.Get(6).Width = 77.0!
        NumberCellType4.AllowEditorVerticalAlign = True
        NumberCellType4.DecimalPlaces = 0
        NumberCellType4.FixedPoint = False
        NumberCellType4.MaximumValue = 9999999999.0R
        NumberCellType4.MinimumValue = -9999999999.0R
        NumberCellType4.ShowSeparator = True
        Me.fbFutai_Sheet1.Columns.Get(7).CellType = NumberCellType4
        Me.fbFutai_Sheet1.Columns.Get(7).Label = "消費税"
        Me.fbFutai_Sheet1.Columns.Get(7).Width = 77.0!
        Me.fbFutai_Sheet1.Columns.Get(8).Label = "入力摘要1"
        Me.fbFutai_Sheet1.Columns.Get(8).Width = 215.0!
        Me.fbFutai_Sheet1.Columns.Get(9).Label = "入力摘要2"
        Me.fbFutai_Sheet1.Columns.Get(9).Width = 215.0!
        Me.fbFutai_Sheet1.Columns.Get(10).BackColor = System.Drawing.Color.FromArgb(CType(CType(246, Byte), Integer), CType(CType(234, Byte), Integer), CType(CType(210, Byte), Integer))
        Me.fbFutai_Sheet1.Columns.Get(10).Label = "プロジェクト"
        Me.fbFutai_Sheet1.Columns.Get(10).Width = 71.0!
        Me.fbFutai_Sheet1.Columns.Get(11).BackColor = System.Drawing.Color.FromArgb(CType(CType(246, Byte), Integer), CType(CType(234, Byte), Integer), CType(CType(210, Byte), Integer))
        Me.fbFutai_Sheet1.Columns.Get(11).Label = "プロジェクト" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "内訳"
        Me.fbFutai_Sheet1.Columns.Get(11).Width = 117.0!
        ButtonCellType2.ButtonColor2 = System.Drawing.SystemColors.ButtonFace
        ButtonCellType2.Text = "下にコピー"
        Me.fbFutai_Sheet1.Columns.Get(12).CellType = ButtonCellType2
        Me.fbFutai_Sheet1.Columns.Get(12).Label = "コピー"
        Me.fbFutai_Sheet1.Columns.Get(12).Width = 61.0!
        ButtonCellType3.ButtonColor2 = System.Drawing.SystemColors.ButtonFace
        ButtonCellType3.Text = "表示"
        Me.fbFutai_Sheet1.Columns.Get(13).CellType = ButtonCellType3
        Me.fbFutai_Sheet1.Columns.Get(13).Label = "プロジェクト" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "一覧"
        Me.fbFutai_Sheet1.Columns.Get(13).Width = 62.0!
        Me.fbFutai_Sheet1.RowHeader.Columns.Default.Resizable = False
        Me.fbFutai_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.A1
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label1.Location = New System.Drawing.Point(14, 152)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(80, 21)
        Me.Label1.TabIndex = 116
        Me.Label1.Text = "タイトル１　＊"
        '
        'txtTitle1
        '
        Me.txtTitle1.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txtTitle1.Location = New System.Drawing.Point(100, 149)
        Me.txtTitle1.MaxLength = 25
        Me.txtTitle1.Name = "txtTitle1"
        Me.txtTitle1.Size = New System.Drawing.Size(278, 20)
        Me.txtTitle1.TabIndex = 117
        '
        'txtTitle2
        '
        Me.txtTitle2.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.txtTitle2.Location = New System.Drawing.Point(100, 176)
        Me.txtTitle2.MaxLength = 25
        Me.txtTitle2.Name = "txtTitle2"
        Me.txtTitle2.Size = New System.Drawing.Size(278, 20)
        Me.txtTitle2.TabIndex = 119
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label2.Location = New System.Drawing.Point(14, 179)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(80, 21)
        Me.Label2.TabIndex = 118
        Me.Label2.Text = "タイトル２　＊"
        '
        'EXTZ0204
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.BackColor = System.Drawing.Color.White
        Me.BackgroundImage = Global.EXTZ.My.Resources.Resources.背景
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(1684, 676)
        Me.Controls.Add(Me.txtTitle2)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtTitle1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.fbFutai)
        Me.Controls.Add(Me.fbRiyoryo)
        Me.Controls.Add(Me.fbBillPay)
        Me.Controls.Add(Me.lblTaxSagaku)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.lblSagakuKin)
        Me.Controls.Add(Me.Label15)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.btnBack)
        Me.Controls.Add(Me.btnComplate)
        Me.Controls.Add(Me.Label65)
        Me.DoubleBuffered = True
        Me.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "EXTZ0204"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ちゃり～ん。　EXASプロジェクト設定"
        CType(Me.fbBillPay, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fbBillPay_Sheet1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fbRiyoryo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fbRiyoryo_Sheet1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fbFutai, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fbFutai_Sheet1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label65 As System.Windows.Forms.Label
    Friend WithEvents btnBack As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents lblTaxSagaku As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents lblSagakuKin As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents btnComplate As System.Windows.Forms.Button
    Friend WithEvents fbBillPay As FarPoint.Win.Spread.FpSpread
    Friend WithEvents fbBillPay_Sheet1 As FarPoint.Win.Spread.SheetView
    Friend WithEvents fbRiyoryo As FarPoint.Win.Spread.FpSpread
    Friend WithEvents fbRiyoryo_Sheet1 As FarPoint.Win.Spread.SheetView
    Friend WithEvents fbFutai As FarPoint.Win.Spread.FpSpread
    Friend WithEvents fbFutai_Sheet1 As FarPoint.Win.Spread.SheetView
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtTitle1 As System.Windows.Forms.TextBox
    Friend WithEvents txtTitle2 As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label

End Class
