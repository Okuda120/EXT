<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class EXTZ0212
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
        Dim CheckBoxCellType2 As FarPoint.Win.Spread.CellType.CheckBoxCellType = New FarPoint.Win.Spread.CellType.CheckBoxCellType()
        Dim CheckBoxCellType3 As FarPoint.Win.Spread.CellType.CheckBoxCellType = New FarPoint.Win.Spread.CellType.CheckBoxCellType()
        Dim CheckBoxCellType4 As FarPoint.Win.Spread.CellType.CheckBoxCellType = New FarPoint.Win.Spread.CellType.CheckBoxCellType()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(EXTZ0212))
        Me.Label65 = New System.Windows.Forms.Label()
        Me.btnBack = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btnKakutei = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.fbRiyobi = New FarPoint.Win.Spread.FpSpread()
        Me.fbRiyobi_Sheet1 = New FarPoint.Win.Spread.SheetView()
        Me.fbBairitu = New FarPoint.Win.Spread.FpSpread()
        Me.fbBairitu_Sheet1 = New FarPoint.Win.Spread.SheetView()
        Me.fbRyokin = New FarPoint.Win.Spread.FpSpread()
        Me.fbRyokin_Sheet1 = New FarPoint.Win.Spread.SheetView()
        Me.fbBunrui = New FarPoint.Win.Spread.FpSpread()
        Me.fbBunrui_Sheet1 = New FarPoint.Win.Spread.SheetView()
        CType(Me.fbRiyobi, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fbRiyobi_Sheet1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fbBairitu, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fbBairitu_Sheet1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fbRyokin, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fbRyokin_Sheet1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fbBunrui, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fbBunrui_Sheet1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label65
        '
        Me.Label65.BackColor = System.Drawing.Color.Transparent
        Me.Label65.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label65.Location = New System.Drawing.Point(14, 29)
        Me.Label65.Name = "Label65"
        Me.Label65.Size = New System.Drawing.Size(83, 18)
        Me.Label65.TabIndex = 25
        Me.Label65.Text = "料金"
        '
        'btnBack
        '
        Me.btnBack.Location = New System.Drawing.Point(343, 781)
        Me.btnBack.Name = "btnBack"
        Me.btnBack.Size = New System.Drawing.Size(96, 23)
        Me.btnBack.TabIndex = 86
        Me.btnBack.Text = "戻る"
        Me.btnBack.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label2.Location = New System.Drawing.Point(70, 646)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(0, 13)
        Me.Label2.TabIndex = 79
        '
        'btnKakutei
        '
        Me.btnKakutei.Location = New System.Drawing.Point(526, 781)
        Me.btnKakutei.Name = "btnKakutei"
        Me.btnKakutei.Size = New System.Drawing.Size(96, 23)
        Me.btnKakutei.TabIndex = 89
        Me.btnKakutei.Text = "確定"
        Me.btnKakutei.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label1.Location = New System.Drawing.Point(14, 314)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(83, 18)
        Me.Label1.TabIndex = 90
        Me.Label1.Text = "倍率"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(14, 570)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(83, 18)
        Me.Label3.TabIndex = 91
        Me.Label3.Text = "反映する日付"
        '
        'fbRiyobi
        '
        Me.fbRiyobi.AccessibleDescription = "FpSpread1, Sheet1, Row 0, Column 0, "
        Me.fbRiyobi.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Never
        Me.fbRiyobi.Location = New System.Drawing.Point(19, 591)
        Me.fbRiyobi.Name = "fbRiyobi"
        Me.fbRiyobi.Sheets.AddRange(New FarPoint.Win.Spread.SheetView() {Me.fbRiyobi_Sheet1})
        Me.fbRiyobi.Size = New System.Drawing.Size(255, 162)
        Me.fbRiyobi.TabIndex = 139
        '
        'fbRiyobi_Sheet1
        '
        Me.fbRiyobi_Sheet1.Reset()
        Me.fbRiyobi_Sheet1.SheetName = "Sheet1"
        'Formulas and custom names must be loaded with R1C1 reference style
        Me.fbRiyobi_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.R1C1
        Me.fbRiyobi_Sheet1.Cells.Get(0, 1).Value = "2015/5/3（日）"
        Me.fbRiyobi_Sheet1.Cells.Get(1, 1).Value = "2015/5/3（月）"
        Me.fbRiyobi_Sheet1.Cells.Get(2, 1).Value = "2015/5/5（火）"
        Me.fbRiyobi_Sheet1.ColumnHeader.Cells.Get(0, 0).Value = "選択"
        Me.fbRiyobi_Sheet1.ColumnHeader.Cells.Get(0, 1).Value = "コピー先の利用日"
        Me.fbRiyobi_Sheet1.Columns.Get(0).CellType = CheckBoxCellType1
        Me.fbRiyobi_Sheet1.Columns.Get(0).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.fbRiyobi_Sheet1.Columns.Get(0).Label = "選択"
        Me.fbRiyobi_Sheet1.Columns.Get(0).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.fbRiyobi_Sheet1.Columns.Get(0).Width = 37.0!
        Me.fbRiyobi_Sheet1.Columns.Get(1).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.fbRiyobi_Sheet1.Columns.Get(1).Label = "コピー先の利用日"
        Me.fbRiyobi_Sheet1.Columns.Get(1).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.fbRiyobi_Sheet1.Columns.Get(1).Width = 158.0!
        Me.fbRiyobi_Sheet1.RowHeader.Columns.Default.Resizable = False
        Me.fbRiyobi_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.A1
        '
        'fbBairitu
        '
        Me.fbBairitu.AccessibleDescription = "FpSpread2, Sheet1, Row 0, Column 0, "
        Me.fbBairitu.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Never
        Me.fbBairitu.Location = New System.Drawing.Point(17, 335)
        Me.fbBairitu.Name = "fbBairitu"
        Me.fbBairitu.Sheets.AddRange(New FarPoint.Win.Spread.SheetView() {Me.fbBairitu_Sheet1})
        Me.fbBairitu.Size = New System.Drawing.Size(381, 206)
        Me.fbBairitu.TabIndex = 140
        '
        'fbBairitu_Sheet1
        '
        Me.fbBairitu_Sheet1.Reset()
        Me.fbBairitu_Sheet1.SheetName = "Sheet1"
        'Formulas and custom names must be loaded with R1C1 reference style
        Me.fbBairitu_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.R1C1
        Me.fbBairitu_Sheet1.ColumnCount = 3
        Me.fbBairitu_Sheet1.RowCount = 0
        Me.fbBairitu_Sheet1.ActiveColumnIndex = -1
        Me.fbBairitu_Sheet1.ActiveRowIndex = -1
        Me.fbBairitu_Sheet1.ColumnHeader.Cells.Get(0, 0).Value = "倍率名"
        Me.fbBairitu_Sheet1.ColumnHeader.Cells.Get(0, 1).Value = "倍率"
        Me.fbBairitu_Sheet1.ColumnHeader.Cells.Get(0, 2).Value = "選択"
        Me.fbBairitu_Sheet1.Columns.Get(0).Label = "倍率名"
        Me.fbBairitu_Sheet1.Columns.Get(0).Width = 216.0!
        Me.fbBairitu_Sheet1.Columns.Get(2).CellType = CheckBoxCellType2
        Me.fbBairitu_Sheet1.Columns.Get(2).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.fbBairitu_Sheet1.Columns.Get(2).Label = "選択"
        Me.fbBairitu_Sheet1.Columns.Get(2).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.fbBairitu_Sheet1.Columns.Get(2).Width = 45.0!
        Me.fbBairitu_Sheet1.RowHeader.Columns.Default.Resizable = False
        Me.fbBairitu_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.A1
        '
        'fbRyokin
        '
        Me.fbRyokin.AccessibleDescription = "FpSpread3, Sheet1, Row 0, Column 0, 月～木（ドリンク無）"
        Me.fbRyokin.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Never
        Me.fbRyokin.Location = New System.Drawing.Point(291, 50)
        Me.fbRyokin.Name = "fbRyokin"
        Me.fbRyokin.Sheets.AddRange(New FarPoint.Win.Spread.SheetView() {Me.fbRyokin_Sheet1})
        Me.fbRyokin.Size = New System.Drawing.Size(602, 240)
        Me.fbRyokin.TabIndex = 141
        '
        'fbRyokin_Sheet1
        '
        Me.fbRyokin_Sheet1.Reset()
        Me.fbRyokin_Sheet1.SheetName = "Sheet1"
        'Formulas and custom names must be loaded with R1C1 reference style
        Me.fbRyokin_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.R1C1
        Me.fbRyokin_Sheet1.ColumnCount = 5
        Me.fbRyokin_Sheet1.RowCount = 0
        Me.fbRyokin_Sheet1.ActiveColumnIndex = -1
        Me.fbRyokin_Sheet1.ActiveRowIndex = -1
        Me.fbRyokin_Sheet1.ColumnHeader.Cells.Get(0, 0).Value = "料金名"
        Me.fbRyokin_Sheet1.ColumnHeader.Cells.Get(0, 1).Value = "時間貸し"
        Me.fbRyokin_Sheet1.ColumnHeader.Cells.Get(0, 2).Value = "選択"
        Me.fbRyokin_Sheet1.ColumnHeader.Cells.Get(0, 3).Value = "利用料金" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "（1日,Lockout）"
        Me.fbRyokin_Sheet1.ColumnHeader.Cells.Get(0, 4).Value = "選択"
        Me.fbRyokin_Sheet1.ColumnHeader.Rows.Get(0).Height = 35.0!
        Me.fbRyokin_Sheet1.Columns.Get(0).Label = "料金名"
        Me.fbRyokin_Sheet1.Columns.Get(0).Width = 215.0!
        Me.fbRyokin_Sheet1.Columns.Get(1).Label = "時間貸し"
        Me.fbRyokin_Sheet1.Columns.Get(1).Width = 122.0!
        Me.fbRyokin_Sheet1.Columns.Get(2).CellType = CheckBoxCellType3
        Me.fbRyokin_Sheet1.Columns.Get(2).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.fbRyokin_Sheet1.Columns.Get(2).Label = "選択"
        Me.fbRyokin_Sheet1.Columns.Get(2).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.fbRyokin_Sheet1.Columns.Get(2).Width = 43.0!
        Me.fbRyokin_Sheet1.Columns.Get(3).Label = "利用料金" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "（1日,Lockout）"
        Me.fbRyokin_Sheet1.Columns.Get(3).Width = 122.0!
        Me.fbRyokin_Sheet1.Columns.Get(4).CellType = CheckBoxCellType4
        Me.fbRyokin_Sheet1.Columns.Get(4).HorizontalAlignment = FarPoint.Win.Spread.CellHorizontalAlignment.Center
        Me.fbRyokin_Sheet1.Columns.Get(4).Label = "選択"
        Me.fbRyokin_Sheet1.Columns.Get(4).VerticalAlignment = FarPoint.Win.Spread.CellVerticalAlignment.Center
        Me.fbRyokin_Sheet1.Columns.Get(4).Width = 43.0!
        Me.fbRyokin_Sheet1.RowHeader.Columns.Default.Resizable = False
        Me.fbRyokin_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.A1
        '
        'fbBunrui
        '
        Me.fbBunrui.AccessibleDescription = "fbBunrui, Sheet1, Row 0, Column 0, 一般　A.着席"
        Me.fbBunrui.HorizontalScrollBarPolicy = FarPoint.Win.Spread.ScrollBarPolicy.Never
        Me.fbBunrui.Location = New System.Drawing.Point(17, 50)
        Me.fbBunrui.Name = "fbBunrui"
        Me.fbBunrui.Sheets.AddRange(New FarPoint.Win.Spread.SheetView() {Me.fbBunrui_Sheet1})
        Me.fbBunrui.Size = New System.Drawing.Size(254, 225)
        Me.fbBunrui.TabIndex = 142
        '
        'fbBunrui_Sheet1
        '
        Me.fbBunrui_Sheet1.Reset()
        Me.fbBunrui_Sheet1.SheetName = "Sheet1"
        'Formulas and custom names must be loaded with R1C1 reference style
        Me.fbBunrui_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.R1C1
        Me.fbBunrui_Sheet1.ColumnCount = 1
        Me.fbBunrui_Sheet1.RowCount = 0
        Me.fbBunrui_Sheet1.ActiveColumnIndex = -1
        Me.fbBunrui_Sheet1.ActiveRowIndex = -1
        Me.fbBunrui_Sheet1.ColumnHeader.Cells.Get(0, 0).Value = "料金分類名"
        Me.fbBunrui_Sheet1.Columns.Get(0).Label = "料金分類名"
        Me.fbBunrui_Sheet1.Columns.Get(0).Width = 197.0!
        Me.fbBunrui_Sheet1.RowHeader.Columns.Default.Resizable = False
        Me.fbBunrui_Sheet1.ReferenceStyle = FarPoint.Win.Spread.Model.ReferenceStyle.A1
        '
        'EXTZ0212
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.BackColor = System.Drawing.Color.White
        Me.BackgroundImage = Global.EXTZ.My.Resources.Resources.背景
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(998, 816)
        Me.Controls.Add(Me.fbBunrui)
        Me.Controls.Add(Me.fbRyokin)
        Me.Controls.Add(Me.fbBairitu)
        Me.Controls.Add(Me.fbRiyobi)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnKakutei)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.btnBack)
        Me.Controls.Add(Me.Label65)
        Me.DoubleBuffered = True
        Me.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "EXTZ0212"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ちゃり～ん。　単価／倍率選択"
        CType(Me.fbRiyobi, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fbRiyobi_Sheet1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fbBairitu, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fbBairitu_Sheet1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fbRyokin, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fbRyokin_Sheet1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fbBunrui, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fbBunrui_Sheet1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label65 As System.Windows.Forms.Label
    Friend WithEvents btnBack As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnKakutei As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents fbRiyobi As FarPoint.Win.Spread.FpSpread
    Friend WithEvents fbRiyobi_Sheet1 As FarPoint.Win.Spread.SheetView
    Friend WithEvents fbBairitu As FarPoint.Win.Spread.FpSpread
    Friend WithEvents fbBairitu_Sheet1 As FarPoint.Win.Spread.SheetView
    Friend WithEvents fbRyokin As FarPoint.Win.Spread.FpSpread
    Friend WithEvents fbRyokin_Sheet1 As FarPoint.Win.Spread.SheetView
    Friend WithEvents fbBunrui As FarPoint.Win.Spread.FpSpread
    Friend WithEvents fbBunrui_Sheet1 As FarPoint.Win.Spread.SheetView

End Class
